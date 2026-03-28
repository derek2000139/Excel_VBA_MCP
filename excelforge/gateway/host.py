from __future__ import annotations

import argparse
import sys
import warnings
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.profile_resolver import BundleRegistry, ProfileResolutionError, ProfileResolver
from excelforge.gateway.runtime_client_manager import get_global_runtime_client
from excelforge.gateway.runtime_identity import (
    get_host_identity,
    resolve_runtime_identity,
)
from excelforge.gateway.tool_manifest_registry import ToolManifestRegistry
from excelforge.gateway.utils import call_runtime


TOOL_MANIFEST_MAP: dict[str, str] = {
    "server.get_status": "server.status",
    "server.health": "server.health",
    "workbook.open_file": "workbook.open_file",
    "workbook.create_file": "workbook.create_file",
    "workbook.save_file": "workbook.save_file",
    "workbook.close_file": "workbook.close_file",
    "workbook.inspect": "workbook.inspect",
    "names.inspect": "names.inspect",
    "names.manage": "names.manage",
    "sheet.create_sheet": "sheet.create_sheet",
    "sheet.rename_sheet": "sheet.rename_sheet",
    "sheet.delete_sheet": "sheet.delete_sheet",
    "sheet.set_auto_filter": "sheet.set_auto_filter",
    "sheet.get_conditional_formats": "sheet.get_conditional_formats",
    "sheet.get_data_validations": "sheet.get_data_validations",
    "range.read_values": "range.read_values",
    "range.write_values": "range.write_values",
    "range.clear_contents": "range.clear_contents",
    "range.copy": "range.copy",
    "range.insert_rows": "range.insert_rows",
    "range.delete_rows": "range.delete_rows",
    "range.insert_columns": "range.insert_columns",
    "range.delete_columns": "range.delete_columns",
    "range.sort_data": "range.sort_data",
    "range.merge": "range.merge",
    "format.set_number_format": "format.set_number_format",
    "format.set_font": "format.set_font",
    "format.set_fill": "format.set_fill",
    "format.set_border": "format.set_border",
    "format.set_alignment": "format.set_alignment",
    "format.set_column_width": "format.set_column_width",
    "format.set_row_height": "format.set_row_height",
    "vba.inspect_project": "vba.inspect_project",
    "vba.scan_code": "vba.scan_code",
    "vba.sync_module": "vba.sync_module",
    "vba.remove_module": "vba.remove_module",
    "vba.execute": "vba.execute",
    "vba.compile": "vba.compile",
    "rollback.manage": "rollback.manage",
    "backups.manage": "backups.manage",
    "snapshot.manage": "snapshot.manage",
    "pq.list_connections": "pq.list_connections",
    "pq.list_queries": "pq.list_queries",
    "pq.get_code": "pq.get_code",
    "pq.update_query": "pq.update_query",
    "pq.refresh": "pq.refresh",
    "audit.list_operations": "audit.list_operations",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-mcp",
        description="ExcelForge Unified MCP Host",
    )
    parser.add_argument(
        "--config",
        help="Path to excel-mcp.yaml (optional, uses runtime-config.yaml by default)",
    )
    parser.add_argument(
        "--profile",
        default="basic_edit",
        help="Profile name (default: basic_edit)",
    )
    parser.add_argument(
        "--enable-bundle",
        action="append",
        default=[],
        dest="enabled_bundles",
        help="Extra bundles to enable (can be repeated)",
    )
    parser.add_argument(
        "--disable-bundle",
        action="append",
        default=[],
        dest="disabled_bundles",
        help="Bundles to disable (can be repeated)",
    )
    parser.add_argument(
        "--strict-profile",
        action="store_true",
        help="Fail immediately if profile not found",
    )
    parser.add_argument(
        "--list-profiles",
        action="store_true",
        help="List available profiles and exit",
    )
    parser.add_argument(
        "--list-bundles",
        action="store_true",
        help="List available bundles and exit",
    )
    parser.add_argument(
        "--runtime-scope",
        default="default",
        help="Runtime scope (default: default)",
    )
    parser.add_argument(
        "--runtime-instance",
        default="default",
        help="Runtime instance name (default: default)",
    )
    parser.add_argument(
        "--print-runtime-endpoint",
        action="store_true",
        help="Print resolved Runtime endpoint on startup",
    )
    return parser


def list_profiles_and_exit(profiles_path: Path | None = None) -> None:
    resolver = ProfileResolver(profiles_path)
    profiles = resolver.list_profiles()
    print("Available profiles:")
    for name in profiles:
        info = resolver.get_profile_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    bundles: {', '.join(info['bundles'])}")


def list_bundles_and_exit(bundles_path: Path | None = None) -> None:
    registry = BundleRegistry(bundles_path)
    bundles = registry.list_bundles()
    print("Available bundles:")
    for name in bundles:
        info = registry.get_bundle_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    domains: {', '.join(info['domains'])}")


def check_tool_budget(tool_count: int, profile_info: dict[str, Any]) -> None:
    budget = profile_info.get("tool_budget")
    if budget is None:
        return
    if tool_count > budget:
        warnings.warn(
            f"Tool count ({tool_count}) exceeds budget ({budget}) for profile '{profile_info['name']}'. "
            f"Consider reducing enabled bundles.",
            UserWarning,
        )


def create_host_runtime_client(args: argparse.Namespace) -> Any:
    scope = args.runtime_scope
    instance_name = args.runtime_instance
    runtime_data_dir = None
    if args.config:
        config_path = Path(args.config)
        runtime_config_path = str(config_path.resolve())
    else:
        runtime_config_path = "./runtime-config.yaml"

    identity = resolve_runtime_identity(
        runtime_data_dir=runtime_data_dir,
        scope=scope,
        instance_name=instance_name,
    )
    client = get_global_runtime_client(
        identity=identity,
        auto_start=True,
        connect_timeout=10,
        call_timeout=30,
        runtime_config_path=runtime_config_path,
    )
    return client


def register_tools_for_profile(
    mcp: FastMCP,
    runtime: Any,
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)

    check_tool_budget(len(enabled_tools), profile_info)

    for tool_name in enabled_tools:
        runtime_method = TOOL_MANIFEST_MAP.get(tool_name, tool_name)

        def create_handler(runtime_client, tool, method):
            def handler(**kwargs):
                return call_runtime(
                    runtime_client,
                    tool_name=tool,
                    method=method,
                    params=kwargs,
                )
            return handler

        mcp.add_tool(tool_name, tool_name, create_handler(runtime, tool_name, runtime_method))


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    profiles_path = Path(__file__).parent / "profiles.yaml"
    bundles_path = Path(__file__).parent / "bundles.yaml"

    if args.list_profiles:
        list_profiles_and_exit(profiles_path)
        return 0

    if args.list_bundles:
        list_bundles_and_exit(bundles_path)
        return 0

    if args.strict_profile:
        resolver = ProfileResolver(profiles_path)
        try:
            resolver.resolve(args.profile)
        except ProfileResolutionError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 1

    try:
        runtime = create_host_runtime_client(args)
    except Exception as exc:
        print(f"Error creating Runtime client: {exc}", file=sys.stderr)
        return 1

    if args.print_runtime_endpoint:
        identity = get_host_identity()
        print(f"Runtime endpoint: {identity.pipe_name}")
        print(f"Runtime instance ID: {identity.instance_id}")

    display_name = f"ExcelForge ({args.profile})"
    mcp = FastMCP(display_name)

    register_tools_for_profile(
        mcp=mcp,
        runtime=runtime,
        profile_name=args.profile,
        extra_bundles=args.enabled_bundles,
        disabled_bundles=args.disabled_bundles,
        profiles_path=profiles_path,
        bundles_path=bundles_path,
    )

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
