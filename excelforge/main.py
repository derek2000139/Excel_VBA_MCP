from __future__ import annotations

import argparse
import json
from typing import Any

from excelforge.config import write_default_config
from excelforge.gateway.core.main import main as core_gateway_main
from excelforge.gateway.pq.main import main as pq_gateway_main
from excelforge.gateway.recovery.main import main as recovery_gateway_main
from excelforge.gateway.vba.main import main as vba_gateway_main
from excelforge.runtime.main import main as runtime_main
from excelforge.server import healthcheck


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge", description="ExcelForge V2 multi-gateway launcher")
    parser.add_argument("--config", default=None, help="Path to runtime config.yaml")

    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("runtime", help="Run Runtime JSON-RPC pipe server")
    core = sub.add_parser("gateway-core", help="Run Core MCP gateway")
    core.add_argument("--gateway-config", required=True, help="Path to excel-core-mcp.yaml")
    vba = sub.add_parser("gateway-vba", help="Run VBA MCP gateway")
    vba.add_argument("--gateway-config", required=True, help="Path to excel-vba-mcp.yaml")
    recovery = sub.add_parser("gateway-recovery", help="Run Recovery MCP gateway")
    recovery.add_argument("--gateway-config", required=True, help="Path to excel-recovery-mcp.yaml")
    pq = sub.add_parser("gateway-pq", help="Run PQ MCP gateway")
    pq.add_argument("--gateway-config", required=True, help="Path to excel-pq-mcp.yaml")
    sub.add_parser("healthcheck", help="Validate runtime prerequisites")
    sub.add_parser("write-default-config", help="Write a default config.yaml")

    return parser


def _print_json(data: dict[str, Any]) -> None:
    print(json.dumps(data, ensure_ascii=False, indent=2))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "write-default-config":
        path = write_default_config("config.yaml")
        _print_json({"ok": True, "config_path": str(path)})
        return 0

    if args.command == "healthcheck":
        data = healthcheck(config_path=args.config)
        _print_json(data)
        return 0

    if args.command == "runtime":
        return runtime_main(["--config", args.config] if args.config else [])
    if args.command == "gateway-core":
        return core_gateway_main(["--config", args.gateway_config])
    if args.command == "gateway-vba":
        return vba_gateway_main(["--config", args.gateway_config])
    if args.command == "gateway-recovery":
        return recovery_gateway_main(["--config", args.gateway_config])
    if args.command == "gateway-pq":
        return pq_gateway_main(["--config", args.gateway_config])

    parser.error(f"Unknown command: {args.command}")
    return 2
