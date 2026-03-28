from __future__ import annotations

import os
from pathlib import Path
from typing import Any

import yaml
from pydantic import BaseModel, ConfigDict, Field


def _workspace_root() -> Path:
    return Path(__file__).resolve().parent.parent


class ServerConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    version: str = "1.0.1"
    actor_id: str = "local-mcp-client"


class RuntimeConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    version: str = "2.0.0"
    pipe_name: str = r"\\.\pipe\excelforge-runtime"
    data_dir: str = ".runtime_data_v2"


class ExcelConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    visible: bool = False
    disable_events: bool = True
    disable_alerts: bool = True
    force_disable_macros: bool = True
    health_ping_enabled: bool = True
    max_rebuild_attempts: int = Field(default=3, ge=1, le=10)
    ensure_visibility: bool = True
    enable_warmup: bool = True
    startup_timeout_seconds: int = Field(default=120, ge=10, le=600)
    request_wait_ready_seconds: int = Field(default=15, ge=1, le=60)


class PathsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    allowed_roots: list[str]
    allowed_extensions: list[str] = [".xlsx", ".xlsm", ".xlsb", ".xls", ".bas", ".cls"]
    snapshots_dir: str
    backups_dir: str
    sqlite_path: str


class LimitsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    max_open_workbooks: int = 8
    max_read_cells: int = 10_000
    max_write_cells: int = 10_000
    max_copy_cells: int = 10_000
    max_snapshot_cells: int = 10_000
    default_read_rows: int = 200
    max_read_rows: int = 1000
    operation_timeout_seconds: int = 30
    max_create_sheets: int = 20
    max_insert_rows: int = 1000
    max_insert_columns: int = 100
    max_vba_code_size_bytes: int = 1_048_576
    calculation_timeout_seconds: int = 10


class SnapshotConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    max_per_workbook: int = 50
    max_total_size_mb: int = 200
    max_age_hours: int = 24
    cleanup_interval_ops: int = 100
    preview_token_ttl_minutes: int = 5


class BackupConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    max_per_workbook: int = 10
    max_total_size_mb: int = 500
    max_age_hours: int = 48
    confirm_token_ttl_minutes: int = 5


class RetentionConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    audit_days: int = 30


class VbaPolicyConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    enabled: bool = True
    block_levels: list[str] = ["critical", "high"]
    warn_levels: list[str] = ["medium"]
    max_code_size_bytes: int = 1_048_576
    allow_execution: bool = True
    execution_timeout_seconds: int = 30
    max_inline_code_size_bytes: int = 524_288


class ToolsGroupsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    core: bool = True
    vba: bool = True
    recovery: bool = True
    names: bool = True


class ToolsConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile: str = "default"
    visible_tool_budget: int = 40
    strict_budget_enforcement: bool = True
    groups: ToolsGroupsConfig = Field(default_factory=ToolsGroupsConfig)


class AppConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    server: ServerConfig
    runtime: RuntimeConfig = Field(default_factory=RuntimeConfig)
    excel: ExcelConfig
    paths: PathsConfig
    limits: LimitsConfig
    snapshot: SnapshotConfig
    backup: BackupConfig
    retention: RetentionConfig
    vba_policy: VbaPolicyConfig = Field(default_factory=VbaPolicyConfig)
    tools: ToolsConfig = Field(default_factory=ToolsConfig)

    @property
    def allowed_roots(self) -> list[Path]:
        result = []
        for p in self.paths.allowed_roots:
            if p == "*":
                result.append(Path("*"))
            else:
                result.append(Path(p).resolve())
        return result

    @property
    def snapshots_dir(self) -> Path:
        return Path(self.paths.snapshots_dir).resolve()

    @property
    def backups_dir(self) -> Path:
        return Path(self.paths.backups_dir).resolve()

    @property
    def sqlite_path(self) -> Path:
        return Path(self.paths.sqlite_path).resolve()


def _default_config() -> AppConfig:
    root = _workspace_root()
    runtime_cfg = RuntimeConfig()
    data_dir = (root / runtime_cfg.data_dir).resolve()
    return AppConfig(
        server=ServerConfig(),
        runtime=runtime_cfg,
        excel=ExcelConfig(),
        paths=PathsConfig(
            allowed_roots=[str(root)],
            snapshots_dir=str(data_dir / "snapshots"),
            backups_dir=str(data_dir / "backups"),
            sqlite_path=str(data_dir / "excelforge.db"),
        ),
        limits=LimitsConfig(),
        snapshot=SnapshotConfig(),
        backup=BackupConfig(),
        retention=RetentionConfig(),
    )


def _deep_merge(base: dict[str, Any], updates: dict[str, Any]) -> dict[str, Any]:
    for key, value in updates.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            base[key] = _deep_merge(base[key], value)
        else:
            base[key] = value
    return base


def _parse_env_value(raw: str) -> Any:
    lowered = raw.lower()
    if lowered in {"true", "false"}:
        return lowered == "true"
    if lowered in {"null", "none"}:
        return None
    if raw.isdigit() or (raw.startswith("-") and raw[1:].isdigit()):
        return int(raw)
    if "," in raw:
        return [part.strip() for part in raw.split(",") if part.strip()]
    return raw


def _env_overrides(prefix: str = "EXCELFORGE_") -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in os.environ.items():
        if not key.startswith(prefix):
            continue
        suffix = key[len(prefix) :]
        if suffix in {"CONFIG"}:
            continue
        parts = [p.lower() for p in suffix.split("__") if p]
        if not parts:
            continue
        cursor = result
        for part in parts[:-1]:
            if part not in cursor or not isinstance(cursor[part], dict):
                cursor[part] = {}
            cursor = cursor[part]
        cursor[parts[-1]] = _parse_env_value(value)
    return result


def ensure_runtime_dirs(cfg: AppConfig) -> None:
    cfg.sqlite_path.parent.mkdir(parents=True, exist_ok=True)
    cfg.snapshots_dir.mkdir(parents=True, exist_ok=True)
    cfg.backups_dir.mkdir(parents=True, exist_ok=True)


def load_config(path: str | Path | None = None) -> AppConfig:
    defaults = _default_config()
    config_path = Path(path) if path else Path(os.environ.get("EXCELFORGE_CONFIG", "config.yaml"))
    merged: dict[str, Any] = defaults.model_dump()

    if config_path.exists():
        with config_path.open("r", encoding="utf-8") as f:
            file_data = yaml.safe_load(f) or {}
        if not isinstance(file_data, dict):
            raise ValueError("config.yaml must contain a top-level mapping")
        _deep_merge(merged, file_data)

    env_data = _env_overrides()
    _deep_merge(merged, env_data)

    cfg = AppConfig.model_validate(merged)
    ensure_runtime_dirs(cfg)
    return cfg


def write_default_config(path: str | Path = "config.yaml") -> Path:
    cfg = _default_config().model_dump(mode="json")
    dest = Path(path)
    with dest.open("w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
    return dest
