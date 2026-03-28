from __future__ import annotations

from typing import Any

from excelforge.models.audit_models import AuditListOperationsRequest
from excelforge.runtime_api.context import RuntimeApiContext


class AuditApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def list_operations(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AuditListOperationsRequest(
            workbook_id=params.get("workbook_id"),
            tool_name=params.get("tool_name"),
            success_only=bool(params.get("success_only", False)),
            limit=int(params.get("limit", 20)),
            offset=int(params.get("offset", 0)),
            operation_id=params.get("operation_id"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="audit.list_operations",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.audit_service.list_operations(
                workbook_id=req.workbook_id,
                tool_name=req.tool_name,
                success_only=req.success_only,
                limit=req.limit,
                offset=req.offset,
                operation_id=req.operation_id,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "tool_name": req.tool_name,
                "success_only": req.success_only,
                "limit": req.limit,
                "offset": req.offset,
                "operation_id": req.operation_id,
            },
            default_workbook_id=req.workbook_id,
        )
