from __future__ import annotations

import threading
import time
from collections.abc import Callable
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.common import ToolEnvelope, error_envelope, ok_envelope
from excelforge.models.error_models import ErrorCode, ExcelForgeError, normalize_exception
from excelforge.persistence.cleanup import CleanupService
from excelforge.services.audit_service import AuditContext, AuditService
from excelforge.utils.ids import generate_id
from excelforge.utils.timestamps import utc_now_rfc3339


class OperationService:
    def __init__(
        self,
        config: AppConfig,
        audit_service: AuditService,
        cleanup_service: CleanupService,
    ) -> None:
        self._config = config
        self._audit_service = audit_service
        self._cleanup_service = cleanup_service
        self._counter = 0
        self._lock = threading.Lock()

    def run_cleanup_on_startup(self) -> None:
        self._cleanup_service.run()

    def run(
        self,
        *,
        tool_name: str,
        operation_fn: Callable[[], dict[str, Any]],
        client_request_id: str | None,
        args_summary: dict[str, Any] | None = None,
        default_workbook_id: str | None = None,
        default_file_path: str | None = None,
        client_name: str | None = None,
        actor_id: str | None = None,
    ) -> ToolEnvelope:
        operation_id = generate_id("op")
        started_at = utc_now_rfc3339()
        t0 = time.perf_counter()
        success = False
        code: str = ErrorCode.OK.value
        message = "operation completed"
        data: dict[str, Any] | None = None

        workbook_id_for_meta = default_workbook_id
        snapshot_id_for_meta: str | None = None
        affected_sheet: str | None = None
        affected_range: str | None = None
        rollback_supported = False
        warnings: list[str] = []

        try:
            data = operation_fn()
            success = True
            if isinstance(data, dict):
                workbook_id_for_meta = cast_str(data.get("workbook_id"), workbook_id_for_meta)
                snapshot_id_for_meta = cast_str(data.get("snapshot_id"), None)
                affected_sheet = cast_str(data.get("sheet_name"), None)
                affected_range = cast_str(
                    data.get("affected_range") or data.get("restored_range") or data.get("range"),
                    None,
                )
                rollback_supported = snapshot_id_for_meta is not None
                raw_warnings = data.get("__warnings__")
                if isinstance(raw_warnings, list):
                    warnings = [str(item) for item in raw_warnings]
                    data.pop("__warnings__", None)
        except Exception as exc:  # noqa: BLE001
            err = normalize_exception(exc)
            code = err.code.value
            message = err.message
            success = False

        duration_ms = int((time.perf_counter() - t0) * 1000)

        self._audit_service.record_operation(
            AuditContext(
                tool_name=tool_name,
                operation_id=operation_id,
                workbook_id=workbook_id_for_meta,
                file_path=default_file_path,
                started_at=started_at,
                duration_ms=duration_ms,
                success=success,
                code=code,
                message=message,
                affected_sheet=affected_sheet,
                affected_range=affected_range,
                snapshot_id=snapshot_id_for_meta,
                args_summary=args_summary,
                client_request_id=client_request_id,
                client_name=client_name,
                actor_id=actor_id,
            )
        )

        self._maybe_periodic_cleanup()

        if success:
            return ok_envelope(
                tool_name=tool_name,
                operation_id=operation_id,
                duration_ms=duration_ms,
                server_version=self._config.server.version,
                data=data,
                workbook_id=workbook_id_for_meta,
                snapshot_id=snapshot_id_for_meta,
                rollback_supported=rollback_supported,
                client_request_id=client_request_id,
                warnings=warnings,
            )

        return error_envelope(
            tool_name=tool_name,
            operation_id=operation_id,
            duration_ms=duration_ms,
            server_version=self._config.server.version,
            code=code,
            message=message,
            workbook_id=workbook_id_for_meta,
            snapshot_id=snapshot_id_for_meta,
            client_request_id=client_request_id,
            warnings=warnings,
        )

    def _maybe_periodic_cleanup(self) -> None:
        interval = int(self._config.snapshot.cleanup_interval_ops)
        if interval <= 0:
            return
        with self._lock:
            self._counter += 1
            if self._counter % interval != 0:
                return
        self._cleanup_service.run()


def cast_str(value: Any, fallback: str | None) -> str | None:
    if value is None:
        return fallback
    return str(value)
