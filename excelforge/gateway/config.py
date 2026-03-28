from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import BaseModel, ConfigDict, Field


class GatewaySection(BaseModel):
    model_config = ConfigDict(extra="forbid")

    id: str
    display_name: str
    runtime_data_dir: str
    auto_start_runtime: bool = True
    runtime_config_path: str | None = None
    connect_timeout_seconds: int = Field(default=10, ge=1, le=60)
    call_timeout_seconds: int = Field(default=30, ge=1, le=600)


class GatewayConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    gateway: GatewaySection


def load_gateway_config(path: str | Path) -> GatewayConfig:
    cfg_path = Path(path)
    if not cfg_path.exists():
        raise ValueError(f"Gateway config not found: {cfg_path}")
    payload = yaml.safe_load(cfg_path.read_text(encoding="utf-8")) or {}
    if not isinstance(payload, dict):
        raise ValueError("Gateway config must be a top-level mapping")
    return GatewayConfig.model_validate(payload)
