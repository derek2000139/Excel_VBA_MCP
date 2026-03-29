from __future__ import annotations

from typing import Any

from excelforge.models.table_models import (
    TableCreateRequest,
    TableDeleteRequest,
    TableInspectRequest,
    TableListRequest,
    TableRenameRequest,
    TableResizeRequest,
    TableSetStyleRequest,
    TableToggleTotalRowRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class TableApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def list_tables(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableListRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.list_tables",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.list_tables(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def create(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableCreateRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range_address=params.get("range_address", ""),
            table_name=params.get("table_name"),
            has_header=params.get("has_header", True),
            style_name=params.get("style_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.create",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.create_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range_address": req.range_address,
                "table_name": req.table_name,
                "has_header": req.has_header,
                "style_name": req.style_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def inspect(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableInspectRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.inspect",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.inspect_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def resize(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableResizeRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            new_range_address=params.get("new_range_address", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.resize",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.resize_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "new_range_address": req.new_range_address,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def rename(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableRenameRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            new_name=params.get("new_name", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.rename",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.rename_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "new_name": req.new_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def set_style(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableSetStyleRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            style_name=params.get("style_name"),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.set_style",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.set_table_style(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "style_name": req.style_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def toggle_total_row(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableToggleTotalRowRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.toggle_total_row",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.toggle_total_row(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def delete(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = TableDeleteRequest(
            workbook_id=params.get("workbook_id", ""),
            table_name=params.get("table_name", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="table.delete",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.table_service.delete_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
