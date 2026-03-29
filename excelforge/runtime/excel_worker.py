from __future__ import annotations

import queue
import threading
from concurrent.futures import Future, TimeoutError as FutureTimeoutError
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Generic, TypeVar

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_app import ExcelAppManager
from excelforge.runtime.retry_policy import run_with_com_retry
from excelforge.runtime.workbook_registry import WorkbookRegistry
from excelforge.runtime.worker_health import WorkerHealth, WorkerMetrics
from excelforge.runtime.worker_manager import ExcelWorkerManager
from excelforge.utils.ids import compute_runtime_fingerprint
from excelforge.utils.timestamps import utc_now_rfc3339

T = TypeVar("T")


@dataclass
class WorkerContext:
    app_manager: ExcelAppManager
    registry: WorkbookRegistry
    worker: ExcelWorker


@dataclass
class WorkerTask(Generic[T]):
    func: Callable[[WorkerContext], T]
    future: Future[T]
    allow_rebuild: bool
    requires_excel: bool


class ExcelWorker:
    def __init__(self, config: AppConfig) -> None:
        self._config = config
        self._queue: queue.Queue[WorkerTask[Any] | None] = queue.Queue()
        self._state = "stopped"
        self._hard_stopped = False
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._last_health_ping: str | None = None
        self._rebuild_count = 0
        self._last_rebuild_at: str | None = None
        runtime_fingerprint = compute_runtime_fingerprint(
            config.runtime.pipe_name,
            str(Path(config.runtime.data_dir).resolve()),
        )
        self._context = WorkerContext(
            app_manager=ExcelAppManager(config),
            registry=WorkbookRegistry(runtime_fingerprint=runtime_fingerprint),
            worker=self,
        )
        self._ready_event = threading.Event()
        self._warmup_started = False
        self._warmup_error: Exception | None = None
        self._excel_version: str | None = None
        self._worker_manager = ExcelWorkerManager()
        self._metrics = WorkerMetrics()

    @property
    def state(self) -> str:
        return self._state

    @property
    def metrics(self) -> WorkerMetrics:
        return self._metrics

    @property
    def queue_length(self) -> int:
        return self._queue.qsize()

    @property
    def context(self) -> WorkerContext:
        return self._context

    @property
    def generation(self) -> int:
        return self._context.registry.generation

    @property
    def last_health_ping(self) -> str | None:
        return self._last_health_ping

    @property
    def rebuild_count(self) -> int:
        return self._rebuild_count

    @property
    def last_rebuild_at(self) -> str | None:
        return self._last_rebuild_at

    def get_metrics(self) -> dict:
        return self._metrics.to_dict()

    def record_operation(self, high_risk: bool = False) -> None:
        self._metrics.record_operation(high_risk)

    def record_exception(self, exc_type: str) -> None:
        self._metrics.record_exception(exc_type)

    def get_excel_pid(self) -> int | None:
        return self._metrics.excel_pid

    def start(self) -> None:
        with self._lock:
            if self._thread and self._thread.is_alive():
                return
            self._state = "running"
            self._hard_stopped = False
            self._thread = threading.Thread(target=self._run_loop, name="excel-sta-worker", daemon=True)
            self._thread.start()

        import logging
        logger = logging.getLogger(__name__)
        try:
            cleaned = self._worker_manager.scan_and_cleanup_orphans()
            if cleaned > 0:
                logger.info(f"[ExcelWorker] Cleaned {cleaned} orphan(s) on startup")
        except Exception as e:
            logger.warning(f"[ExcelWorker] Startup orphan scan failed: {e}")

    def submit(
        self,
        func: Callable[[WorkerContext], T],
        *,
        timeout_seconds: int,
        allow_rebuild: bool = False,
        requires_excel: bool = True,
    ) -> T:
        current_state = self.state
        if requires_excel and current_state == "degraded":
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_RECOVERING,
                "Excel worker is recovering, please retry shortly",
            )
        if requires_excel and current_state == "stopped" and self._hard_stopped:
            raise ExcelForgeError(
                ErrorCode.E500_EXCEL_UNAVAILABLE,
                "Excel worker is stopped after repeated rebuild failures",
            )

        if current_state == "stopped":
            self.start()

        fut: Future[T] = Future()
        task = WorkerTask(
            func=func,
            future=fut,
            allow_rebuild=allow_rebuild,
            requires_excel=requires_excel,
        )
        self._queue.put(task)

        try:
            return fut.result(timeout=timeout_seconds)
        except FutureTimeoutError as exc:
            raise ExcelForgeError(
                ErrorCode.E500_OPERATION_TIMEOUT,
                f"Operation timed out after {timeout_seconds} seconds",
            ) from exc

    def stop(self, wait_seconds: int = 15) -> None:
        with self._lock:
            thread = self._thread
            if thread is None:
                self._state = "stopped"
                return
            self._state = "stopped"
        self._queue.put(None)
        thread.join(timeout=wait_seconds)

    def _set_state(self, state: str) -> None:
        with self._lock:
            self._state = state

    def _run_loop(self) -> None:
        pythoncom = None
        try:
            import pythoncom as _pythoncom  # type: ignore

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
        except Exception:
            self._set_state("degraded")

        try:
            while True:
                task = self._queue.get()
                if task is None:
                    break

                if task.future.cancelled():
                    continue

                try:
                    result = self._execute_task(task)
                except Exception as exc:  # noqa: BLE001
                    task.future.set_exception(exc)
                else:
                    task.future.set_result(result)
        finally:
            self._context.app_manager.close()
            self._context.registry.invalidate_all()
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    def _execute_task(self, task: WorkerTask[T]) -> T:
        high_risk = task.allow_rebuild
        self.record_operation(high_risk=high_risk)

        if task.requires_excel:
            if self._config.excel.health_ping_enabled and not self._health_ping():
                self._set_state("degraded")
                recovered = self._recover_excel_instance()
                if not recovered:
                    self.record_exception("EXCEL_UNAVAILABLE")
                    raise ExcelForgeError(
                        ErrorCode.E500_EXCEL_UNAVAILABLE,
                        "Excel instance unavailable after rebuild attempts",
                    )
                if not task.allow_rebuild:
                    raise ExcelForgeError(
                        ErrorCode.E410_WORKBOOK_STALE,
                        "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                    )
            else:
                self._context.app_manager.ensure_app()
                if self.state != "running":
                    self._set_state("running")

        try:
            return run_with_com_retry(lambda: task.func(self._context))
        except ExcelForgeError as exc:
            self.record_exception(exc.code.value if exc.code else type(exc).__name__)
            if exc.code != ErrorCode.E500_COM_DISCONNECTED:
                raise

            self._set_state("degraded")
            recovered = self._recover_excel_instance()
            if not recovered:
                raise ExcelForgeError(
                    ErrorCode.E500_EXCEL_UNAVAILABLE,
                    "Excel instance unavailable after rebuild attempts",
                ) from exc

            if not task.allow_rebuild:
                raise ExcelForgeError(
                    ErrorCode.E410_WORKBOOK_STALE,
                    "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                ) from exc

            return run_with_com_retry(lambda: task.func(self._context))

    def _health_ping(self) -> bool:
        self._last_health_ping = utc_now_rfc3339()
        if not self._context.app_manager.ready:
            try:
                app = self._context.app_manager.ensure_app()
                self._update_excel_pid(app)
                return True
            except Exception:
                return False
        if self._context.app_manager.ping():
            app = self._context.app_manager._app
            if app is not None:
                self._update_excel_pid(app)
            return True
        return False

    def _update_excel_pid(self, app: Any) -> None:
        try:
            from excelforge.runtime.com_utils import get_excel_pid
            pid = get_excel_pid(app)
            if pid and pid != self._metrics.excel_pid:
                self._metrics.excel_pid = pid
        except Exception:
            pass

    def _recover_excel_instance(self) -> bool:
        def _create_new_app():
            old_pid = self._metrics.excel_pid
            self._context.app_manager._app = None
            app = self._context.app_manager.ensure_app()
            try:
                from excelforge.runtime.com_utils import get_excel_pid
                pid = get_excel_pid(app)
                self._worker_manager.register_worker_pid(pid)
                if old_pid and pid and old_pid != pid:
                    self._metrics.last_pid_change_reason = "stale_rebuild"
                self._metrics.excel_pid = pid
            except Exception:
                pass
            return app

        def _pre_rebuild_hook():
            self._context.registry.clear_all()
            self._worker_manager.clear_registration()
            import logging
            logging.getLogger(__name__).warning(
                f"[ExcelWorker] === REBUILD START === rebuild_count={self._rebuild_count}"
            )

        import logging
        logger = logging.getLogger(__name__)

        self._invalidate_excel_session()
        self._warmup_error = None
        self._ready_event.clear()
        attempts = int(self._config.excel.max_rebuild_attempts)

        for attempt in range(attempts):
            try:
                app = self._worker_manager.rebuild_worker(
                    create_fn=_create_new_app,
                    pre_rebuild_hook=_pre_rebuild_hook
                )
                self._excel_version = app.Version
                self._warmup_error = None
            except Exception as exc:
                self._warmup_error = exc
                logger.error(f"[ExcelWorker] Rebuild attempt {attempt + 1} failed: {exc}")
                continue
            self._rebuild_count += 1
            self._last_rebuild_at = utc_now_rfc3339()
            self._metrics.last_recycle_reason = "rebuild_attempt"
            self._hard_stopped = False
            self._set_state("running")
            self._ready_event.set()
            logger.info(f"[ExcelWorker] === REBUILD COMPLETE === count={self._rebuild_count}")
            return True

        self._hard_stopped = True
        self._set_state("stopped")
        self._ready_event.set()
        logger.error("[ExcelWorker] All rebuild attempts failed")
        return False

    def _invalidate_excel_session(self) -> None:
        self._context.registry.bump_generation()
        self._context.app_manager.invalidate()

    def warmup(self, timeout_seconds: int = 120) -> bool:
        if self._warmup_started and self._ready_event.is_set():
            return self._warmup_error is None

        if not self._warmup_started:
            self._warmup_started = True
            self.start()

            def _do_warmup(ctx: WorkerContext) -> None:
                try:
                    app = ctx.app_manager.ensure_app()
                    self._excel_version = app.Version
                    self._warmup_error = None
                    try:
                        from excelforge.runtime.com_utils import get_excel_pid
                        pid = get_excel_pid(app)
                        if pid:
                            self._worker_manager.register_worker_pid(pid)
                            self._metrics.excel_pid = pid
                    except Exception:
                        pass
                except Exception as exc:
                    self._warmup_error = exc
                finally:
                    self._ready_event.set()

            fut: Future[None] = Future()
            task = WorkerTask(
                func=_do_warmup,
                future=fut,
                allow_rebuild=False,
                requires_excel=True,
            )
            self._queue.put(task)

        finished = self._ready_event.wait(timeout=timeout_seconds)
        if not finished:
            return False
        return self._warmup_error is None

    def is_ready(self) -> bool:
        return self._ready_event.is_set() and self._warmup_error is None

    def wait_ready(self, timeout: float | None = None) -> bool:
        finished = self._ready_event.wait(timeout=timeout)
        if not finished:
            return False
        return self._warmup_error is None

    def get_ready_status(self) -> dict:
        return {
            "ready": self.is_ready(),
            "version": self._excel_version,
            "warmup_started": self._warmup_started,
            "warmup_error": str(self._warmup_error) if self._warmup_error else None,
        }

    def mark_unhealthy(self, reason: str) -> None:
        self._set_state("degraded")
        self._recover_excel_instance()

    def rebuild(self, reopen_workbooks: bool = False, reason: str = "manual_rebuild") -> dict:
        import logging
        logger = logging.getLogger(__name__)

        old_pid = self._metrics.excel_pid
        open_workbook_paths: list[str] = []
        for handle in self._context.registry.list_items():
            open_workbook_paths.append(handle.file_path)

        if reopen_workbooks and open_workbook_paths:
            logger.info(f"[ExcelWorker] rebuild will reopen {len(open_workbook_paths)} workbook(s) after rebuild")
        else:
            logger.info(f"[ExcelWorker] rebuild will NOT auto-reopen workbooks (reopen_workbooks={reopen_workbooks})")

        self._metrics.last_recycle_reason = reason
        self._metrics.state = WorkerHealth.RECYCLING

        self._invalidate_excel_session()

        try:
            self._context.app_manager.close()
        except Exception as e:
            logger.warning(f"[ExcelWorker] Error closing app during rebuild: {e}")

        self._invalidate_excel_session()

        def _create_new_app():
            old_pid = self._metrics.excel_pid
            self._context.app_manager._app = None
            app = self._context.app_manager.ensure_app()
            try:
                from excelforge.runtime.com_utils import get_excel_pid
                pid = get_excel_pid(app)
                self._worker_manager.register_worker_pid(pid)
                if old_pid and pid and old_pid != pid:
                    self._metrics.last_pid_change_reason = reason
                self._metrics.excel_pid = pid
            except Exception:
                pass
            return app

        def _pre_rebuild_hook():
            self._context.registry.clear_all()
            self._worker_manager.clear_registration()
            logger.warning(f"[ExcelWorker] === REBUILD START (rebuild) === reason={reason}")

        self._warmup_error = None
        self._ready_event.clear()
        attempts = int(self._config.excel.max_rebuild_attempts)

        for attempt in range(attempts):
            try:
                app = self._worker_manager.rebuild_worker(
                    create_fn=_create_new_app,
                    pre_rebuild_hook=_pre_rebuild_hook
                )
                self._excel_version = app.Version
                self._warmup_error = None
            except Exception as exc:
                self._warmup_error = exc
                logger.error(f"[ExcelWorker] Rebuild attempt {attempt + 1} failed: {exc}")
                continue
            self._rebuild_count += 1
            self._last_rebuild_at = utc_now_rfc3339()
            self._metrics.last_recycle_reason = reason
            self._metrics.state = WorkerHealth.HEALTHY
            self._hard_stopped = False
            self._set_state("running")
            self._ready_event.set()
            logger.info(f"[ExcelWorker] === REBUILD COMPLETE === count={self._rebuild_count}, reason={reason}")
            return {
                "success": True,
                "rebuild_count": self._rebuild_count,
                "excel_pid": self._metrics.excel_pid,
                "open_workbook_paths": open_workbook_paths if reopen_workbooks else [],
                "reopened": reopen_workbooks,
            }

        self._hard_stopped = True
        self._set_state("stopped")
        self._metrics.state = WorkerHealth.FAILED
        self._ready_event.set()
        logger.error("[ExcelWorker] All rebuild attempts failed")
        return {
            "success": False,
            "error": "All rebuild attempts failed",
        }
