from __future__ import annotations

import json
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
from pathlib import Path
from typing import Any

from excelforge.gateway.runtime_identity import (
    RuntimeIdentity,
    get_host_identity,
    identity_to_dict,
)
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.lifecycle import (
    RuntimeLockInfo,
    is_process_alive,
    read_runtime_lock_from_dir,
)


_GLOBAL_RUNTIME_CLIENT: RuntimeClientManager | None = None
_GLOBAL_CLIENT_LOCK = __import__("threading").Lock()


def get_global_runtime_client(
    identity: RuntimeIdentity | None = None,
    auto_start: bool = True,
    connect_timeout: int = 10,
    call_timeout: int = 30,
    runtime_config_path: str | None = None,
) -> RuntimeClientManager:
    with _GLOBAL_CLIENT_LOCK:
        global _GLOBAL_RUNTIME_CLIENT
        if _GLOBAL_RUNTIME_CLIENT is None:
            if identity is None:
                identity = get_host_identity()
            _GLOBAL_RUNTIME_CLIENT = RuntimeClientManager(
                identity=identity,
                auto_start=auto_start,
                connect_timeout=connect_timeout,
                call_timeout=call_timeout,
                runtime_config_path=runtime_config_path,
            )
        return _GLOBAL_RUNTIME_CLIENT


def reset_global_runtime_client() -> None:
    with _GLOBAL_CLIENT_LOCK:
        global _GLOBAL_RUNTIME_CLIENT
        if _GLOBAL_RUNTIME_CLIENT is not None:
            _GLOBAL_RUNTIME_CLIENT.close()
            _GLOBAL_RUNTIME_CLIENT = None


class RuntimeClientManager:
    def __init__(
        self,
        identity: RuntimeIdentity,
        auto_start: bool = True,
        connect_timeout: int = 10,
        call_timeout: int = 30,
        runtime_config_path: str | None = None,
    ) -> None:
        self._identity = identity
        self._auto_start = auto_start
        self._connect_timeout = connect_timeout
        self._call_timeout = call_timeout
        self._runtime_config_path = runtime_config_path
        self._request_id = 0
        self._pipe = None
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="runtime-client")
        self._connected_identity: str | None = None

    @property
    def identity(self) -> RuntimeIdentity:
        return self._identity

    def close(self) -> None:
        if self._pipe is None:
            return
        try:
            import win32file
            win32file.CloseHandle(self._pipe)
        except Exception:
            pass
        finally:
            self._pipe = None
            self._connected_identity = None

    def call(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        payload = dict(params)
        payload["actor_id"] = self._identity.instance_id

        try:
            return self._call_with_timeout(method, payload)
        except ExcelForgeError as exc:
            if exc.code != ErrorCode.E503_RUNTIME_UNAVAILABLE:
                raise
            self.close()
            return self._call_with_timeout(method, payload)

    def _call_with_timeout(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        fut = self._executor.submit(self._call_blocking, method, params)
        try:
            return fut.result(timeout=self._call_timeout)
        except FutureTimeoutError as exc:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_TIMEOUT,
                f"Runtime call timed out after {self._call_timeout} seconds",
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
            import pywintypes
            import win32file
        except Exception as exc:
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
        if self._pipe is not None and self._connected_identity == self._identity.instance_id:
            return
        lock = self._wait_for_runtime_lock()
        if lock is None:
            if self._auto_start:
                self._start_runtime_process()
                lock = self._wait_for_runtime_lock()
            if lock is None:
                raise ExcelForgeError(
                    ErrorCode.E503_RUNTIME_UNAVAILABLE,
                    f"Runtime lock file not found or process is not alive for identity {self._identity.instance_id}",
                )
        self._connect_pipe(lock.pipe_name)
        self._connected_identity = self._identity.instance_id

    def _wait_for_runtime_lock(self) -> RuntimeLockInfo | None:
        start = time.time()
        while time.time() - start <= self._connect_timeout:
            lock = read_runtime_lock_from_dir(str(self._identity.data_dir))
            if lock is not None and is_process_alive(lock.pid):
                if lock.pipe_name == self._identity.pipe_name:
                    return lock
            time.sleep(0.1)
        return None

    def _start_runtime_process(self) -> None:
        if not self._runtime_config_path:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_UNAVAILABLE,
                "runtime_config_path is required when auto_start_runtime=true",
            )
        cmd = [
            sys.executable,
            "-m",
            "excelforge.runtime",
            "--config",
            str(Path(self._runtime_config_path).resolve()),
        ]
        subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)

    def _connect_pipe(self, pipe_name: str) -> None:
        try:
            import pywintypes
            import win32file
            import win32pipe
        except Exception as exc:
            raise ExcelForgeError(ErrorCode.E503_RUNTIME_UNAVAILABLE, f"pywin32 unavailable: {exc}") from exc

        start = time.time()
        while time.time() - start <= self._connect_timeout:
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

    def get_runtime_info(self) -> dict[str, Any]:
        return identity_to_dict(self._identity)

    def get_connection_status(self) -> dict[str, Any]:
        lock = read_runtime_lock_from_dir(str(self._identity.data_dir))
        return {
            "connected": self._pipe is None,
            "instance_id": self._identity.instance_id,
            "pipe_name": self._identity.pipe_name,
            "lock_exists": lock is not None,
            "lock_pid": lock.pid if lock else None,
            "process_alive": is_process_alive(lock.pid) if lock else False,
        }
