from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from excelforge.models.common import ToolEnvelope
from excelforge.runtime.bootstrap import RuntimeServices


@dataclass
class RuntimeApiContext:
    services: RuntimeServices

    def run_operation(
        self,
        *,
        method_name: str,
        actor_id: str,
        client_request_id: str | None,
        operation_fn: Callable[[], dict[str, Any]],
        args_summary: dict[str, Any] | None = None,
        default_workbook_id: str | None = None,
        default_file_path: str | None = None,
    ) -> dict[str, Any]:
        envelope: ToolEnvelope = self.services.operation_service.run(
            tool_name=method_name,
            operation_fn=operation_fn,
            client_request_id=client_request_id,
            args_summary=args_summary,
            default_workbook_id=default_workbook_id,
            default_file_path=default_file_path,
            client_name=actor_id,
            actor_id=actor_id,
        )
        return envelope.model_dump(mode="json")
