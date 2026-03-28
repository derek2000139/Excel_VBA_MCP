from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.config import load_gateway_config
from excelforge.gateway.runtime_client import RuntimeClient
from excelforge.gateway.utils import call_runtime


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.gateway.vba", description="ExcelForge VBA MCP Gateway")
    parser.add_argument("--config", required=True, help="Path to excel-vba-mcp.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = load_gateway_config(args.config)
    runtime = RuntimeClient(config)
    mcp = FastMCP(config.gateway.display_name or "ExcelForge VBA")

    @mcp.tool(name="vba.inspect_project")
    def vba_inspect_project(workbook_id: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="vba.inspect_project",
            method="vba.inspect_project",
            params={"workbook_id": workbook_id, "client_request_id": client_request_id},
        )

    @mcp.tool(name="vba.scan_code")
    def vba_scan_code(
        code: str,
        module_name: str = "",
        module_type: str = "standard_module",
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            runtime,
            tool_name="vba.scan_code",
            method="vba.scan_code",
            params={
                "code": code,
                "module_name": module_name,
                "module_type": module_type,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="vba.sync_module")
    def vba_sync_module(
        workbook_id: str,
        module_name: str,
        module_type: str = "standard_module",
        code: str = "",
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            runtime,
            tool_name="vba.sync_module",
            method="vba.sync_module",
            params={
                "workbook_id": workbook_id,
                "module_name": module_name,
                "module_type": module_type,
                "code": code,
                "overwrite": overwrite,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="vba.remove_module")
    def vba_remove_module(workbook_id: str, module_name: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="vba.remove_module",
            method="vba.remove_module",
            params={
                "workbook_id": workbook_id,
                "module_name": module_name,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="vba.execute")
    def vba_execute(
        action: str,
        workbook_id: str,
        procedure_name: str = "",
        arguments: list | None = None,
        code: str = "",
        timeout_seconds: int = 30,
        client_request_id: str = "",
    ) -> dict:
        if action == "macro":
            method = "vba.execute_macro"
            payload = {
                "workbook_id": workbook_id,
                "procedure_name": procedure_name,
                "arguments": arguments or [],
                "timeout_seconds": timeout_seconds,
                "client_request_id": client_request_id,
            }
        else:
            method = "vba.execute_inline"
            payload = {
                "workbook_id": workbook_id,
                "code": code,
                "procedure_name": procedure_name or "Main",
                "timeout_seconds": timeout_seconds,
                "client_request_id": client_request_id,
            }
        return call_runtime(runtime, tool_name="vba.execute", method=method, params=payload)

    @mcp.tool(name="vba.compile")
    def vba_compile(workbook_id: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="vba.compile",
            method="vba.compile",
            params={"workbook_id": workbook_id, "client_request_id": client_request_id},
        )

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
