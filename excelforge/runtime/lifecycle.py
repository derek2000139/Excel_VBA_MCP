from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.utils.timestamps import utc_now_rfc3339


@dataclass
class RuntimeLockInfo:
    pid: int
    pipe_name: str
    started_at: str
    version: str
    config_path: str


def runtime_lock_path(config: AppConfig) -> Path:
    data_dir = Path(config.runtime.data_dir).resolve()
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir / "runtime.lock"


def write_runtime_lock(config: AppConfig, config_path: str | None = None) -> RuntimeLockInfo:
    lock = RuntimeLockInfo(
        pid=os.getpid(),
        pipe_name=config.runtime.pipe_name,
        started_at=utc_now_rfc3339(),
        version=config.runtime.version,
        config_path=str(Path(config_path).resolve()) if config_path else "",
    )
    path = runtime_lock_path(config)
    payload: dict[str, Any] = {
        "pid": lock.pid,
        "pipe_name": lock.pipe_name,
        "started_at": lock.started_at,
        "version": lock.version,
        "config_path": lock.config_path,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return lock


def remove_runtime_lock(config: AppConfig) -> None:
    path = runtime_lock_path(config)
    path.unlink(missing_ok=True)


def is_process_alive(pid: int) -> bool:
    """
    检查指定 PID 的进程是否存在。

    Windows 兼容实现：不使用 os.kill(pid, 0)，
    因为该调用在 Windows 上行为不一致，
    某些 Python 版本下会产生 C 层异常状态污染。
    """
    if pid <= 0:
        return False

    if sys.platform == "win32":
        return _is_process_alive_windows(pid)
    else:
        return _is_process_alive_posix(pid)


def _is_process_alive_windows(pid: int) -> bool:
    """Windows 实现：使用 ctypes 调用 OpenProcess + GetExitCodeProcess。"""
    import ctypes
    import ctypes.wintypes

    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    STILL_ACTIVE = 259

    kernel32 = ctypes.windll.kernel32

    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if handle == 0:
        return False

    try:
        exit_code = ctypes.wintypes.DWORD()
        if kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            return exit_code.value == STILL_ACTIVE
        else:
            return True
    finally:
        kernel32.CloseHandle(handle)


def _is_process_alive_posix(pid: int) -> bool:
    """POSIX 实现：使用 os.kill(pid, 0)。"""
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False


def read_runtime_lock_from_dir(runtime_data_dir: str) -> RuntimeLockInfo | None:
    path = Path(runtime_data_dir).resolve() / "runtime.lock"
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        lock = RuntimeLockInfo(
            pid=int(payload["pid"]),
            pipe_name=str(payload["pipe_name"]),
            started_at=str(payload.get("started_at", "")),
            version=str(payload.get("version", "")),
            config_path=str(payload.get("config_path", "")),
        )
    except Exception:
        return None
    if not is_process_alive(lock.pid):
        return None
    return lock
