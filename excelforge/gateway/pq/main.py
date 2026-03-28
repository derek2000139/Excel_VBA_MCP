from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.config import load_gateway_config
from excelforge.gateway.runtime_client import RuntimeClient
from excelforge.gateway.utils import call_runtime


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.gateway.pq", description="ExcelForge PQ MCP Gateway")
    parser.add_argument("--config", required=True, help="Path to excel-pq-mcp.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config = load_gateway_config(args.config)
    runtime = RuntimeClient(config)
    mcp = FastMCP(config.gateway.display_name or "ExcelForge PQ")

    @mcp.tool(name="pq.list_queries")
    def pq_list_queries(workbook_id: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="pq.list_queries",
            method="pq.list_queries",
            params={"workbook_id": workbook_id, "client_request_id": client_request_id},
        )

    @mcp.tool(name="pq.get_code")
    def pq_get_code(workbook_id: str, query_name: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="pq.get_code",
            method="pq.get_query_code",
            params={"workbook_id": workbook_id, "query_name": query_name, "client_request_id": client_request_id},
        )

    @mcp.tool(name="pq.update_query")
    def pq_update_query(workbook_id: str, query_name: str, code: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="pq.update_query",
            method="pq.update_query",
            params={
                "workbook_id": workbook_id,
                "query_name": query_name,
                "code": code,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="pq.refresh")
    def pq_refresh(
        workbook_id: str,
        query_name: str = "",
        timeout_seconds: int = 30,
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            runtime,
            tool_name="pq.refresh",
            method="pq.refresh",
            params={
                "workbook_id": workbook_id,
                "query_name": query_name,
                "timeout_seconds": timeout_seconds,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="pq.list_connections")
    def pq_list_connections(workbook_id: str, client_request_id: str = "") -> dict:
        return call_runtime(
            runtime,
            tool_name="pq.list_connections",
            method="pq.list_connections",
            params={"workbook_id": workbook_id, "client_request_id": client_request_id},
        )

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
