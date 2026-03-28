from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_formula_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="formula.fill_range")
    def formula_fill_range(
        workbook_id: str,
        sheet_name: str,
        range: str,
        formula: str,
        formula_type: str = "standard",
        preview_rows: int = 5,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="formula.fill_range",
            method="formula.fill",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "formula": formula,
                "formula_type": formula_type,
                "preview_rows": preview_rows,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="formula.set_single")
    def formula_set_single(
        workbook_id: str,
        sheet_name: str,
        cell: str,
        formula: str,
        formula_type: str = "standard",
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="formula.set_single",
            method="formula.set_single",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "cell": cell,
                "formula": formula,
                "formula_type": formula_type,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="formula.get_dependencies")
    def formula_get_dependencies(
        workbook_id: str,
        sheet_name: str,
        cell: str,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="formula.get_dependencies",
            method="formula.get_dependencies",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "cell": cell,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="formula.repair_references")
    def formula_repair_references(
        workbook_id: str,
        sheet_name: str,
        range: str,
        action: str,
        replacements: list[dict] | None = None,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="formula.repair_references",
            method="formula.repair",
            params={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range,
                "action": action,
                "replacements": replacements,
                "client_request_id": client_request_id,
            },
        )
