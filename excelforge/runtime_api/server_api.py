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
        return self._ctx.run_operation(
            method_name="server.health",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: {
                "runtime_version": self._ctx.services.config.runtime.version,
                "worker_state": self._ctx.services.worker.state,
                "pipe_name": self._ctx.services.config.runtime.pipe_name,
                "excel_ready": self._ctx.services.worker.context.app_manager.ready,
            },
            args_summary={},
        )
