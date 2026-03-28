from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_sheet_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="sheet.inspect_structure")
    def sheet_inspect_structure(
        workbook_id: str,
        sheet_name: str,
        sample_rows: int = 5,
        scan_rows: int = 10,
        max_profile_columns: int = 50,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.inspect_structure",
            method="sheet.inspect",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "sample_rows": sample_rows,
                "scan_rows": scan_rows,
                "max_profile_columns": max_profile_columns,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.create_sheet")
    def sheet_create_sheet(
        workbook_id: str,
        sheet_name: str,
        position: str = "last",
        reference_sheet: str = "",
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.create_sheet",
            method="sheet.create",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "position": position,
                "reference_sheet": reference_sheet,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.rename_sheet")
    def sheet_rename_sheet(workbook_id: str, sheet_name: str, new_name: str, client_request_id: str = "") -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.rename_sheet",
            method="sheet.rename",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "new_name": new_name,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.delete_sheet")
    def sheet_delete_sheet(
        workbook_id: str,
        sheet_name: str,
        preview: bool = False,
        confirm_token: str = "",
        client_request_id: str = "",
    ) -> dict:
        method = "sheet.preview_delete" if preview else "sheet.delete"
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.delete_sheet",
            method=method,
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "confirm_token": confirm_token,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.set_auto_filter")
    def sheet_set_auto_filter(
        workbook_id: str,
        sheet_name: str,
        action: str,
        range: str | None = None,
        filters: list[dict] | None = None,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.set_auto_filter",
            method="sheet.auto_filter",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "action": action,
                "range": range,
                "filters": filters,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.get_conditional_formats")
    def sheet_get_conditional_formats(
        workbook_id: str,
        sheet_name: str,
        range: str = "",
        limit: int = 100,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.get_conditional_formats",
            method="sheet.get_conditional_formats",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "limit": limit,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="sheet.get_data_validations")
    def sheet_get_data_validations(
        workbook_id: str,
        sheet_name: str,
        range: str = "",
        limit: int = 100,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="sheet.get_data_validations",
            method="sheet.get_data_validations",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "limit": limit,
                "client_request_id": client_request_id,
            },
        )
