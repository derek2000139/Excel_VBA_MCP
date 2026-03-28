from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_range_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="range.read_values")
    def range_read_values(
        workbook_id: str,
        sheet_name: str,
        range: str,
        value_mode: str = "raw",
        include_formulas: bool = False,
        row_offset: int = 0,
        row_limit: int = 200,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.read_values",
            method="range.read",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "value_mode": value_mode,
                "include_formulas": include_formulas,
                "row_offset": row_offset,
                "row_limit": row_limit,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.write_values")
    def range_write_values(
        workbook_id: str,
        sheet_name: str,
        start_cell: str,
        values: list[list],
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.write_values",
            method="range.write",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "start_cell": start_cell,
                "values": values,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.clear_contents")
    def range_clear_contents(
        workbook_id: str,
        sheet_name: str,
        range: str,
        scope: str = "contents",
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.clear_contents",
            method="range.clear",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "scope": scope,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.copy")
    def range_copy(
        workbook_id: str,
        sheet_name: str,
        source_range: str,
        target_sheet: str,
        target_start_cell: str,
        paste_mode: str = "values",
        target_workbook_id: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.copy",
            method="range.copy",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "source_range": source_range,
                "target_sheet": target_sheet,
                "target_start_cell": target_start_cell,
                "paste_mode": paste_mode,
                "target_workbook_id": target_workbook_id,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.insert_rows")
    def range_insert_rows(
        workbook_id: str,
        sheet_name: str,
        row: int,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.insert_rows",
            method="range.insert_rows",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "row": row,
                "count": count,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.delete_rows")
    def range_delete_rows(
        workbook_id: str,
        sheet_name: str,
        row: int,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.delete_rows",
            method="range.delete_rows",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "row": row,
                "count": count,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.insert_columns")
    def range_insert_columns(
        workbook_id: str,
        sheet_name: str,
        column: str,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.insert_columns",
            method="range.insert_columns",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "column": column,
                "count": count,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.delete_columns")
    def range_delete_columns(
        workbook_id: str,
        sheet_name: str,
        column: str,
        count: int = 1,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.delete_columns",
            method="range.delete_columns",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "column": column,
                "count": count,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.sort_data")
    def range_sort_data(
        workbook_id: str,
        sheet_name: str,
        range: str,
        sort_fields: list[dict],
        has_header: bool = False,
        case_sensitive: bool = False,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.sort_data",
            method="range.sort",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "sort_fields": sort_fields,
                "has_header": has_header,
                "case_sensitive": case_sensitive,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.merge")
    def range_merge(
        workbook_id: str,
        sheet_name: str,
        range: str,
        across: bool = False,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.merge",
            method="range.merge",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "across": across,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="range.unmerge")
    def range_unmerge(
        workbook_id: str,
        sheet_name: str,
        range: str,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="range.unmerge",
            method="range.unmerge",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "client_request_id": client_request_id,
            },
        )
