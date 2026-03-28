from __future__ import annotations

import os
from pathlib import Path

import yaml

from excelforge.config import load_config
from excelforge.runtime.lifecycle import read_runtime_lock_from_dir, remove_runtime_lock, write_runtime_lock


def test_runtime_lock_write_read_remove(tmp_path: Path) -> None:
    config_path = tmp_path / "runtime-config.yaml"
    data_dir = tmp_path / ".runtime_data_v2"
    cfg = {
        "server": {"version": "2.0.0", "actor_id": "runtime"},
        "runtime": {"version": "2.0.0", "pipe_name": r"\\.\pipe\excelforge-runtime-test", "data_dir": str(data_dir)},
        "excel": {
            "visible": False,
            "disable_events": True,
            "disable_alerts": True,
            "force_disable_macros": True,
            "health_ping_enabled": True,
            "max_rebuild_attempts": 3,
            "ensure_visibility": True,
        },
        "paths": {
            "allowed_roots": [str(tmp_path)],
            "snapshots_dir": str(data_dir / "snapshots"),
            "backups_dir": str(data_dir / "backups"),
            "sqlite_path": str(data_dir / "excelforge.db"),
        },
        "limits": {},
        "snapshot": {},
        "backup": {},
        "retention": {},
    }
    config_path.write_text(yaml.safe_dump(cfg), encoding="utf-8")
    app_cfg = load_config(config_path)

    lock = write_runtime_lock(app_cfg, str(config_path))
    assert lock.pid == os.getpid()

    parsed = read_runtime_lock_from_dir(str(data_dir))
    assert parsed is not None
    assert parsed.pipe_name == r"\\.\pipe\excelforge-runtime-test"

    remove_runtime_lock(app_cfg)
    assert read_runtime_lock_from_dir(str(data_dir)) is None
