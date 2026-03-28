from __future__ import annotations

from pathlib import Path
from typing import Any

from excelforge.gateway.runtime_client_manager import get_global_runtime_client
from excelforge.gateway.runtime_identity import resolve_runtime_identity


ENTRYPOINT_DEFAULT_PROFILE = {
    "excel-core-mcp": "legacy_core",
    "excel-vba-mcp": "legacy_vba",
    "excel-recovery-mcp": "legacy_recovery",
    "excel-pq-mcp": "legacy_pq",
}


def create_legacy_runtime_client(entrypoint_id: str) -> Any:
    identity = resolve_runtime_identity(
        runtime_data_dir=None,
        scope="default",
        instance_name="default",
    )
    client = get_global_runtime_client(
        identity=identity,
        auto_start=True,
        connect_timeout=10,
        call_timeout=30,
        runtime_config_path="./runtime-config.yaml",
    )
    return client


def get_legacy_profile(entrypoint_id: str) -> str:
    return ENTRYPOINT_DEFAULT_PROFILE.get(entrypoint_id, "legacy_core")
