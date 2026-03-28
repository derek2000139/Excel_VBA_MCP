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

        return {
            "success": True,
            "code": "OK",
            "message": message,
            "data": {
                "runtime_status": "running",
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
