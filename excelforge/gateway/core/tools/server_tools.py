from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_server_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="server.get_status")
    def server_get_status(client_request_id: str = "") -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="server.get_status",
            method="server.status",
            params={"client_request_id": client_request_id},
        )
