from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.utils import call_runtime


def register_workbook_tools(mcp: FastMCP, ctx: GatewayToolContext) -> None:
    @mcp.tool(name="workbook.open_file")
    def workbook_open_file(file_path: str, read_only: bool = False, client_request_id: str = "") -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="workbook.open_file",
            method="workbook.open",
            params={
                "file_path": file_path,
                "read_only": read_only,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="workbook.inspect")
    def workbook_inspect(action: str, workbook_id: str = "", client_request_id: str = "") -> dict:
        method = "workbook.list" if action == "list" else "workbook.info"
        return call_runtime(
            ctx.runtime,
            tool_name="workbook.inspect",
            method=method,
            params={"workbook_id": workbook_id, "client_request_id": client_request_id},
        )

    @mcp.tool(name="workbook.save_file")
    def workbook_save_file(workbook_id: str, save_as_path: str = "", client_request_id: str = "") -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="workbook.save_file",
            method="workbook.save",
            params={
                "workbook_id": workbook_id,
                "save_as_path": save_as_path,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="workbook.close_file")
    def workbook_close_file(
        workbook_id: str,
        force_discard: bool = False,
        client_request_id: str = "",
    ) -> dict:
        save_before_close = not force_discard
        return call_runtime(
            ctx.runtime,
            tool_name="workbook.close_file",
            method="workbook.close",
            params={
                "workbook_id": workbook_id,
                "save_before_close": save_before_close,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="workbook.create_file")
    def workbook_create_file(
        file_path: str,
        sheet_names: list[str] | None = None,
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            ctx.runtime,
            tool_name="workbook.create_file",
            method="workbook.create",
            params={
                "file_path": file_path,
                "sheet_names": sheet_names or ["Sheet1"],
                "overwrite": overwrite,
                "client_request_id": client_request_id,
            },
        )
