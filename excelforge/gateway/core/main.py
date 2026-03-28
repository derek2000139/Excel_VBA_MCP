from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.core.tools import (
    register_format_tools,
    register_formula_tools,
    register_range_tools,
    register_server_tools,
    register_sheet_tools,
    register_workbook_tools,
)
from excelforge.gateway.core.tools.common import GatewayToolContext
from excelforge.gateway.legacy_wrapper import create_legacy_runtime_client


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.gateway.core", description="ExcelForge Core MCP Gateway")
    parser.add_argument("--config", required=True, help="Path to excel-core-mcp.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    runtime = create_legacy_runtime_client("excel-core-mcp")
    mcp = FastMCP("ExcelForge Core (Legacy)")
    ctx = GatewayToolContext(runtime=runtime)

    register_server_tools(mcp, ctx)
    register_workbook_tools(mcp, ctx)
    register_sheet_tools(mcp, ctx)
    register_range_tools(mcp, ctx)
    register_formula_tools(mcp, ctx)
    register_format_tools(mcp, ctx)

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
