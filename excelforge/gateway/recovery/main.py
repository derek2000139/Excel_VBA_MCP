from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP

from excelforge.gateway.legacy_wrapper import create_legacy_runtime_client
from excelforge.gateway.utils import call_runtime


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.gateway.recovery", description="ExcelForge Recovery MCP Gateway")
    parser.add_argument("--config", required=True, help="Path to excel-recovery-mcp.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    runtime = create_legacy_runtime_client("excel-recovery-mcp")
    mcp = FastMCP("ExcelForge Recovery (Legacy)")

    @mcp.tool(name="rollback.manage")
    def rollback_manage(
        action: str,
        workbook_id: str = "",
        snapshot_id: str = "",
        preview_token: str = "",
        sample_limit: int = 20,
        limit: int = 20,
        offset: int = 0,
        max_age_hours: int | None = None,
        dry_run: bool = False,
        client_request_id: str = "",
    ) -> dict:
        if action == "list":
            method = "recovery.list_snapshots"
            payload = {"workbook_id": workbook_id or None, "limit": limit, "offset": offset}
        elif action == "preview":
            method = "recovery.preview_snapshot"
            payload = {"snapshot_id": snapshot_id, "sample_limit": sample_limit}
        elif action == "restore":
            method = "recovery.restore_snapshot"
            payload = {"snapshot_id": snapshot_id, "preview_token": preview_token}
        elif action == "undo_last":
            method = "recovery.undo_last"
            payload = {"workbook_id": workbook_id}
        elif action == "stats":
            method = "recovery.snapshot_stats"
            payload = {"workbook_id": workbook_id or None}
        elif action == "cleanup":
            method = "recovery.snapshot_cleanup"
            payload = {"workbook_id": workbook_id or None, "max_age_hours": max_age_hours, "dry_run": dry_run}
        else:
            return {
                "success": False,
                "code": "E400_INVALID_ARGUMENT",
                "message": f"Unsupported rollback action: {action}",
                "data": None,
                "meta": {
                    "tool_name": "rollback.manage",
                    "operation_id": "op_gateway",
                    "duration_ms": 0,
                    "server_version": "2.0.0",
                    "workbook_id": workbook_id or None,
                    "snapshot_id": None,
                    "rollback_supported": False,
                    "client_request_id": client_request_id,
                    "warnings": [],
                },
            }
        payload["client_request_id"] = client_request_id
        return call_runtime(runtime, tool_name="rollback.manage", method=method, params=payload)

    @mcp.tool(name="backups.manage")
    def backups_manage(
        action: str,
        workbook_id: str = "",
        backup_id: str = "",
        limit: int = 20,
        offset: int = 0,
        client_request_id: str = "",
    ) -> dict:
        if action == "list":
            method = "recovery.list_backups"
            payload = {"workbook_id": workbook_id or None, "limit": limit, "offset": offset}
        elif action == "restore":
            method = "recovery.restore_backup"
            payload = {"workbook_id": workbook_id, "backup_id": backup_id}
        else:
            return {
                "success": False,
                "code": "E400_INVALID_ARGUMENT",
                "message": f"Unsupported backup action: {action}",
                "data": None,
                "meta": {
                    "tool_name": "backups.manage",
                    "operation_id": "op_gateway",
                    "duration_ms": 0,
                    "server_version": "2.0.0",
                    "workbook_id": workbook_id or None,
                    "snapshot_id": None,
                    "rollback_supported": False,
                    "client_request_id": client_request_id,
                    "warnings": [],
                },
            }
        payload["client_request_id"] = client_request_id
        return call_runtime(runtime, tool_name="backups.manage", method=method, params=payload)

    @mcp.tool(name="audit.list_operations")
    def audit_list_operations(
        workbook_id: str = "",
        tool_name: str = "",
        success_only: bool = False,
        limit: int = 20,
        offset: int = 0,
        operation_id: str = "",
        client_request_id: str = "",
    ) -> dict:
        return call_runtime(
            runtime,
            tool_name="audit.list_operations",
            method="audit.list_operations",
            params={
                "workbook_id": workbook_id or None,
                "tool_name": tool_name or None,
                "success_only": success_only,
                "limit": limit,
                "offset": offset,
                "operation_id": operation_id or None,
                "client_request_id": client_request_id,
            },
        )

    @mcp.tool(name="names.inspect")
    def names_inspect(
        action: str,
        workbook_id: str,
        range_name: str = "",
        scope: str = "all",
        sheet_name: str = "",
        value_mode: str = "raw",
        row_offset: int = 0,
        row_limit: int = 200,
        client_request_id: str = "",
    ) -> dict:
        if action == "list":
            method = "names.list"
            payload = {
                "workbook_id": workbook_id,
                "scope": scope,
                "sheet_name": sheet_name or None,
            }
        elif action == "info":
            method = "names.read"
            payload = {
                "workbook_id": workbook_id,
                "range_name": range_name,
                "value_mode": value_mode,
                "row_offset": row_offset,
                "row_limit": row_limit,
            }
        else:
            return {
                "success": False,
                "code": "E400_INVALID_ARGUMENT",
                "message": f"Unsupported names.inspect action: {action}",
                "data": None,
                "meta": {
                    "tool_name": "names.inspect",
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
        payload["client_request_id"] = client_request_id
        return call_runtime(runtime, tool_name="names.inspect", method=method, params=payload)

    @mcp.tool(name="names.manage")
    def names_manage(
        action: str,
        workbook_id: str,
        name: str = "",
        refers_to: str = "",
        scope: str = "workbook",
        sheet_name: str = "",
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        if action == "create":
            method = "names.create"
            payload = {
                "workbook_id": workbook_id,
                "name": name,
                "refers_to": refers_to,
                "scope": scope,
                "sheet_name": sheet_name or None,
                "overwrite": overwrite,
            }
        elif action == "delete":
            method = "names.delete"
            payload = {
                "workbook_id": workbook_id,
                "name": name,
                "scope": scope,
                "sheet_name": sheet_name or None,
            }
        else:
            return {
                "success": False,
                "code": "E400_INVALID_ARGUMENT",
                "message": f"Unsupported names.manage action: {action}",
                "data": None,
                "meta": {
                    "tool_name": "names.manage",
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
        payload["client_request_id"] = client_request_id
        return call_runtime(runtime, tool_name="names.manage", method=method, params=payload)

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
