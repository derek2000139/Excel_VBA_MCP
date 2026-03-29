from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

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
from excelforge.tools.registry import ToolRegistry


def register_table_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="table.list_tables")
    def table_list_tables(
        workbook_id: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableListRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.list_tables",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.list_tables(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.list_tables", "table_tools", "table")

    @mcp.tool(name="table.create")
    def table_create(
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        table_name: str | None = None,
        has_header: bool = True,
        style_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableCreateRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            range_address=range_address,
            table_name=table_name,
            has_header=has_header,
            style_name=style_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.create",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.create_table(params=req),
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
        return envelope.model_dump(mode="json")

    registry.add("table.create", "table_tools", "table")

    @mcp.tool(name="table.inspect")
    def table_inspect(
        workbook_id: str,
        table_name: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableInspectRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.inspect",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.inspect_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.inspect", "table_tools", "table")

    @mcp.tool(name="table.resize")
    def table_resize(
        workbook_id: str,
        table_name: str,
        new_range_address: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableResizeRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            new_range_address=new_range_address,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.resize",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.resize_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "new_range_address": req.new_range_address,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.resize", "table_tools", "table")

    @mcp.tool(name="table.rename")
    def table_rename(
        workbook_id: str,
        table_name: str,
        new_name: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableRenameRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            new_name=new_name,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.rename",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.rename_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "new_name": req.new_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.rename", "table_tools", "table")

    @mcp.tool(name="table.set_style")
    def table_set_style(
        workbook_id: str,
        table_name: str,
        style_name: str | None = None,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableSetStyleRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            style_name=style_name,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.set_style",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.set_table_style(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "style_name": req.style_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.set_style", "table_tools", "table")

    @mcp.tool(name="table.toggle_total_row")
    def table_toggle_total_row(
        workbook_id: str,
        table_name: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableToggleTotalRowRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.toggle_total_row",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.toggle_total_row(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.toggle_total_row", "table_tools", "table")

    @mcp.tool(name="table.delete")
    def table_delete(
        workbook_id: str,
        table_name: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = TableDeleteRequest(
            workbook_id=workbook_id,
            table_name=table_name,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="table.delete",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.table_service.delete_table(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "table_name": req.table_name,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("table.delete", "table_tools", "table")
