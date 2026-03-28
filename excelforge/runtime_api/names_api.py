from __future__ import annotations

from typing import Any

from excelforge.models.named_range_models import (
    NamedRangeCreateRangeRequest,
    NamedRangeDeleteRangeRequest,
    NamedRangeListRangesRequest,
    NamedRangeReadValuesRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class NamesApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def list(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = NamedRangeListRangesRequest(
            workbook_id=params.get("workbook_id", ""),
            scope=params.get("scope", "all"),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="names.list",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.named_range_service.list_ranges(
                workbook_id=req.workbook_id,
                scope=req.scope,
                sheet_name=req.sheet_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "scope": req.scope, "sheet_name": req.sheet_name},
            default_workbook_id=req.workbook_id,
        )

    def read(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = NamedRangeReadValuesRequest(
            workbook_id=params.get("workbook_id", ""),
            range_name=params.get("range_name", ""),
            value_mode=params.get("value_mode", "raw"),
            row_offset=int(params.get("row_offset", 0)),
            row_limit=int(params.get("row_limit", 200)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="names.read",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.named_range_service.read_values(
                workbook_id=req.workbook_id,
                range_name=req.range_name,
                value_mode=req.value_mode,
                row_offset=req.row_offset,
                row_limit=req.row_limit,
            ),
            args_summary={"workbook_id": req.workbook_id, "range_name": req.range_name},
            default_workbook_id=req.workbook_id,
        )

    def create(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = NamedRangeCreateRangeRequest(
            workbook_id=params.get("workbook_id", ""),
            name=params.get("name", ""),
            refers_to=params.get("refers_to", ""),
            scope=params.get("scope", "workbook"),
            sheet_name=params.get("sheet_name"),
            overwrite=bool(params.get("overwrite", False)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="names.create",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.named_range_service.create_range(
                workbook_id=req.workbook_id,
                name=req.name,
                refers_to=req.refers_to,
                scope=req.scope,
                sheet_name=req.sheet_name,
                overwrite=req.overwrite,
            ),
            args_summary={"workbook_id": req.workbook_id, "name": req.name, "scope": req.scope},
            default_workbook_id=req.workbook_id,
        )

    def delete(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = NamedRangeDeleteRangeRequest(
            workbook_id=params.get("workbook_id", ""),
            name=params.get("name", ""),
            scope=params.get("scope", "workbook"),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="names.delete",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.named_range_service.delete_range(
                workbook_id=req.workbook_id,
                name=req.name,
                scope=req.scope,
                sheet_name=req.sheet_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "name": req.name, "scope": req.scope},
            default_workbook_id=req.workbook_id,
        )
