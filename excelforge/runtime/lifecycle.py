from __future__ import annotations

import json
import os
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
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
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
