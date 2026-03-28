from __future__ import annotations

import hashlib
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PRODUCT_NAME = "ExcelForge"
PRODUCT_PROTOCOL = "excelforge-runtime"


@dataclass(frozen=True)
class RuntimeIdentity:
    instance_id: str
    pipe_name: str
    lock_file_path: str
    pid_file_path: str
    metadata_path: str
    data_dir: Path
    scope: str
    instance_name: str


def resolve_runtime_identity(
    runtime_data_dir: str | Path | None = None,
    scope: str = "default",
    instance_name: str = "default",
    product_name: str = PRODUCT_NAME,
) -> RuntimeIdentity:
    if runtime_data_dir is None:
        runtime_data_dir = os.environ.get(
            "EXCELFORGE_RUNTIME_DATA_DIR",
            "./.runtime_data_v2",
        )

    data_dir = Path(runtime_data_dir).resolve()
    data_dir.mkdir(parents=True, exist_ok=True)

    scope = _normalize_scope(scope)
    instance_name = _normalize_instance_name(instance_name)

    instance_id = _compute_instance_id(product_name, scope, instance_name)

    pipe_name = rf"\\.\pipe\{product_protocol(scope, instance_name)}"

    lock_file_path = str(data_dir / "runtime.lock")
    pid_file_path = str(data_dir / "runtime.pid")
    metadata_path = str(data_dir / "runtime.metadata.json")

    return RuntimeIdentity(
        instance_id=instance_id,
        pipe_name=pipe_name,
        lock_file_path=lock_file_path,
        pid_file_path=pid_file_path,
        metadata_path=metadata_path,
        data_dir=data_dir,
        scope=scope,
        instance_name=instance_name,
    )


def product_protocol(scope: str, instance_name: str) -> str:
    base = f"{PRODUCT_PROTOCOL}.{scope}"
    if instance_name != "default":
        base = f"{base}.{instance_name}"
    return base


def _normalize_scope(scope: str) -> str:
    scope = scope.strip().lower()
    if not scope:
        scope = "default"
    invalid_chars = set('/\\:*?"<>|')
    scope = "".join(c if c not in invalid_chars else "_" for c in scope)
    return scope or "default"


def _normalize_instance_name(instance_name: str) -> str:
    instance_name = instance_name.strip().lower()
    if not instance_name:
        instance_name = "default"
    invalid_chars = set('/\\:*?"<>|')
    instance_name = "".join(c if c not in invalid_chars else "_" for c in instance_name)
    return instance_name or "default"


def _compute_instance_id(product_name: str, scope: str, instance_name: str) -> str:
    raw = f"{product_name}:{scope}:{instance_name}"
    short_hash = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]
    return f"rt_{short_hash}"


def get_legacy_entrypoint_scope(entrypoint_name: str) -> str:
    entrypoint_to_scope = {
        "excel-core-mcp": "legacy_core",
        "excel-vba-mcp": "legacy_vba",
        "excel-recovery-mcp": "legacy_recovery",
        "excel-pq-mcp": "legacy_pq",
    }
    return entrypoint_to_scope.get(entrypoint_name, "default")


def get_legacy_entrypoint_instance(entrypoint_name: str) -> str:
    return "default"


def identity_to_dict(identity: RuntimeIdentity) -> dict[str, Any]:
    return {
        "instance_id": identity.instance_id,
        "pipe_name": identity.pipe_name,
        "lock_file_path": identity.lock_file_path,
        "pid_file_path": identity.pid_file_path,
        "metadata_path": identity.metadata_path,
        "data_dir": str(identity.data_dir),
        "scope": identity.scope,
        "instance_name": identity.instance_name,
    }


def get_host_identity() -> RuntimeIdentity:
    scope = os.environ.get("EXCELFORGE_RUNTIME_SCOPE", "default")
    instance_name = os.environ.get("EXCELFORGE_RUNTIME_INSTANCE", "default")
    runtime_data_dir = os.environ.get("EXCELFORGE_RUNTIME_DATA_DIR", None)
    return resolve_runtime_identity(
        runtime_data_dir=runtime_data_dir,
        scope=scope,
        instance_name=instance_name,
    )
