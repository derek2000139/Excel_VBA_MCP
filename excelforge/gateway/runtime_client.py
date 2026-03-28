from __future__ import annotations

import json
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
from pathlib import Path
from typing import Any

from excelforge.gateway.config import GatewayConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.lifecycle import read_runtime_lock_from_dir


class RuntimeClient:
    def __init__(self, config: GatewayConfig) -> None:
        self._config = config
        self._request_id = 0
        self._pipe = None
        self._pipe_name: str | None = None
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="runtime-client")

    def close(self) -> None:
        if self._pipe is None:
            return
        try:
            import win32file  # type: ignore

            win32file.CloseHandle(self._pipe)
        except Exception:
            pass
        finally:
            self._pipe = None
            self._pipe_name = None

    def call(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        payload = dict(params)
        payload["actor_id"] = self._config.gateway.id

        try:
            return self._call_with_timeout(method, payload)
        except ExcelForgeError as exc:
            if exc.code != ErrorCode.E503_RUNTIME_UNAVAILABLE:
                raise
            self.close()
            return self._call_with_timeout(method, payload)

    def _call_with_timeout(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        timeout = self._config.gateway.call_timeout_seconds
        fut = self._executor.submit(self._call_blocking, method, params)
        try:
            return fut.result(timeout=timeout)
        except FutureTimeoutError as exc:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_TIMEOUT,
                f"Runtime call timed out after {timeout} seconds",
            ) from exc

    def _call_blocking(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        self._ensure_connected()
        request = {
            "jsonrpc": "2.0",
            "id": self._next_request_id(),
            "method": method,
            "params": params,
        }
        data = (json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8")

        try:
            import pywintypes  # type: ignore
            import win32file  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise ExcelForgeError(ErrorCode.E503_RUNTIME_UNAVAILABLE, f"pywin32 unavailable: {exc}") from exc

        if self._pipe is None:
            raise ExcelForgeError(ErrorCode.E503_RUNTIME_UNAVAILABLE, "Runtime pipe is not connected")

        try:
            win32file.WriteFile(self._pipe, data)
            _, raw = win32file.ReadFile(self._pipe, 1024 * 1024)
        except pywintypes.error as exc:
            self.close()
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_UNAVAILABLE,
                f"Runtime pipe communication failed: {exc}",
            ) from exc

        try:
            response = json.loads(bytes(raw).decode("utf-8"))
        except Exception as exc:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_UNAVAILABLE,
                f"Invalid runtime response payload: {exc}",
            ) from exc

        if "error" in response:
            err = response["error"] or {}
            data_obj = err.get("data") or {}
            error_code = data_obj.get("error_code") or err.get("message") or ErrorCode.E500_INTERNAL.value
            detail = data_obj.get("detail") or err.get("message") or "Runtime call failed"
            try:
                code_enum = ErrorCode(error_code)
            except Exception:
                code_enum = ErrorCode.E500_INTERNAL
            raise ExcelForgeError(code_enum, str(detail))
        return response.get("result") or {}

    def _ensure_connected(self) -> None:
        if self._pipe is not None:
            return
        lock = self._wait_for_runtime_lock()
        if lock is None:
            if self._config.gateway.auto_start_runtime:
                self._start_runtime_process()
                lock = self._wait_for_runtime_lock()
            if lock is None:
                raise ExcelForgeError(
                    ErrorCode.E503_RUNTIME_UNAVAILABLE,
                    "Runtime lock file not found or process is not alive",
                )
        self._connect_pipe(lock.pipe_name)

    def _wait_for_runtime_lock(self):
        timeout = self._config.gateway.connect_timeout_seconds
        start = time.time()
        while time.time() - start <= timeout:
            lock = read_runtime_lock_from_dir(self._config.gateway.runtime_data_dir)
            if lock is not None:
                return lock
            time.sleep(0.1)
        return None

    def _start_runtime_process(self) -> None:
        runtime_config = self._config.gateway.runtime_config_path
        if not runtime_config:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_UNAVAILABLE,
                "runtime_config_path is required when auto_start_runtime=true",
            )
        cmd = [sys.executable, "-m", "excelforge.runtime", "--config", str(Path(runtime_config).resolve())]
        subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)  # type: ignore[attr-defined]

    def _connect_pipe(self, pipe_name: str) -> None:
        try:
            import pywintypes  # type: ignore
            import win32file  # type: ignore
            import win32pipe  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise ExcelForgeError(ErrorCode.E503_RUNTIME_UNAVAILABLE, f"pywin32 unavailable: {exc}") from exc

        timeout = self._config.gateway.connect_timeout_seconds
        start = time.time()
        while time.time() - start <= timeout:
            try:
                pipe = win32file.CreateFile(
                    pipe_name,
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None,
                )
                win32pipe.SetNamedPipeHandleState(pipe, win32pipe.PIPE_READMODE_MESSAGE, None, None)
                self._pipe = pipe
                self._pipe_name = pipe_name
                return
            except pywintypes.error:
                time.sleep(0.1)
        raise ExcelForgeError(
            ErrorCode.E503_RUNTIME_UNAVAILABLE,
            f"Failed to connect to runtime pipe: {pipe_name}",
        )

    def _next_request_id(self) -> str:
        self._request_id += 1
        return f"req_{self._request_id}"
