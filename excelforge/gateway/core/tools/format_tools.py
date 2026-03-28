from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_format_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="format.manage")
    def format_manage(
        action: str,
        workbook_id: str,
        sheet_name: str,
        range: str = "",
        style: dict | None = None,
        columns: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        if action == "set_style":
            return call_runtime(
                ctx.runtime,
                tool_name="format.manage",
                method="format.set_style",
                params={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "range": range,
                    "style": style or {},
                    "client_request_id": client_request_id,
                },
            )
        if action == "auto_fit_columns":
            target_range = columns or range
            return call_runtime(
                ctx.runtime,
                tool_name="format.manage",
                method="format.auto_fit",
                params={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "range": target_range,
                    "client_request_id": client_request_id,
                },
            )
        return {
            "success": False,
            "code": "E400_INVALID_ARGUMENT",
            "message": f"Unsupported action: {action}",
            "data": None,
            "meta": {
                "tool_name": "format.manage",
                "operation_id": "op_gateway",
                "duration_ms": 0,
                "server_version": "2.0.0",
                "workbook_id": workbook_id,
                "snapshot_id": None,
                "rollback_supported": False,
                "client_request_id": client_request_id,
                "warnings": [],
            },
        }
