from __future__ import annotations

from typing import Any

from excelforge.runtime_api.context import RuntimeApiContext


class ServerApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def status(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        _ = params
        _ = actor_id
        return self._ctx.run_operation(
            method_name="server.status",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.server_service.get_status(),
            args_summary={},
        )

    def health(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        _ = params
        _ = actor_id
        ready_status = self._ctx.services.worker.get_ready_status()
        excel_ready = ready_status["ready"]
        warnings: list[str] = []
        message = "Runtime is running"

        if not ready_status["warmup_started"]:
            warnings.append("Excel warmup has not started yet")
        elif not excel_ready:
            if ready_status["warmup_error"]:
                warnings.append(f"Excel initialization failed: {ready_status['warmup_error']}")
            else:
                warnings.append("Excel engine is still initializing")

        runtime_config = self._ctx.services.config.runtime
        data_dir = runtime_config.data_dir
        pipe_name = runtime_config.pipe_name
        instance_id = self._compute_runtime_instance_id(pipe_name, data_dir)

        return {
            "success": True,
            "code": "OK",
            "message": message,
            "data": {
                "runtime_status": "running",
                "runtime_instance_id": instance_id,
                "runtime_endpoint": pipe_name,
                "runtime_pid": self._get_runtime_pid(data_dir),
                "excel": {
                    "ready": excel_ready,
                    "version": ready_status["version"],
                    "warmup_started": ready_status["warmup_started"],
                    "warmup_error": ready_status["warmup_error"],
                },
                "open_workbooks": self._ctx.services.worker.context.registry.count(),
            },
            "warnings": warnings,
            "meta": {},
            "recovery": None,
        }

    def _compute_runtime_instance_id(self, pipe_name: str, data_dir: str) -> str:
        import hashlib
        raw = f"ExcelForge:default:default:{pipe_name}:{data_dir}"
        short_hash = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8]
        return f"rt_{short_hash}"

    def _get_runtime_pid(self, data_dir: str) -> int | None:
        import json
        from pathlib import Path
        lock_path = Path(data_dir).resolve() / "runtime.lock"
        if not lock_path.exists():
            return None
        try:
            payload = json.loads(lock_path.read_text(encoding="utf-8"))
            return int(payload.get("pid", 0)) or None
        except Exception:
            return None
