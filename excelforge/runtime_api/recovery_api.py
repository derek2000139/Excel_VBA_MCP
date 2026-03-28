from __future__ import annotations

from typing import Any

from excelforge.runtime_api.context import RuntimeApiContext


class RecoveryApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def list_snapshots(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        limit = int(params.get("limit", 20))
        offset = int(params.get("offset", 0))
        return self._ctx.run_operation(
            method_name="recovery.list_snapshots",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.rollback_service.list_snapshots(
                workbook_id=workbook_id,
                limit=limit,
                offset=offset,
            ),
            args_summary={"workbook_id": workbook_id, "limit": limit, "offset": offset},
            default_workbook_id=workbook_id,
        )

    def preview_snapshot(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        snapshot_id = str(params.get("snapshot_id", ""))
        sample_limit = int(params.get("sample_limit", 20))
        return self._ctx.run_operation(
            method_name="recovery.preview_snapshot",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.rollback_service.preview_snapshot(
                snapshot_id=snapshot_id,
                sample_limit=sample_limit,
            ),
            args_summary={"snapshot_id": snapshot_id, "sample_limit": sample_limit},
        )

    def restore_snapshot(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        snapshot_id = str(params.get("snapshot_id", ""))
        preview_token = str(params.get("preview_token", ""))
        return self._ctx.run_operation(
            method_name="recovery.restore_snapshot",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.rollback_service.restore_snapshot(
                snapshot_id=snapshot_id,
                preview_token=preview_token,
            ),
            args_summary={"snapshot_id": snapshot_id},
        )

    def undo_last(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        return self._ctx.run_operation(
            method_name="recovery.undo_last",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.rollback_service.undo_last(workbook_id=workbook_id),
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )

    def snapshot_stats(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        return self._ctx.run_operation(
            method_name="recovery.snapshot_stats",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.snapshot_service.get_stats(workbook_id=workbook_id),
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )

    def snapshot_cleanup(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        max_age_hours = params.get("max_age_hours")
        dry_run = bool(params.get("dry_run", False))
        return self._ctx.run_operation(
            method_name="recovery.snapshot_cleanup",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.snapshot_service.run_cleanup(
                max_age_hours=max_age_hours,
                workbook_id=workbook_id,
                dry_run=dry_run,
            ),
            args_summary={"workbook_id": workbook_id, "max_age_hours": max_age_hours, "dry_run": dry_run},
            default_workbook_id=workbook_id,
        )

    def list_backups(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        file_path = params.get("file_path")
        limit = int(params.get("limit", 20))
        offset = int(params.get("offset", 0))
        return self._ctx.run_operation(
            method_name="recovery.list_backups",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.backup_service.list_backups(
                workbook_id=workbook_id,
                file_path=file_path,
                limit=limit,
                offset=offset,
            ),
            args_summary={"workbook_id": workbook_id, "file_path": file_path, "limit": limit, "offset": offset},
            default_workbook_id=workbook_id,
        )

    def restore_backup(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        backup_id = str(params.get("backup_id", ""))
        return self._ctx.run_operation(
            method_name="recovery.restore_backup",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.backup_service.restore_file(
                workbook_id=workbook_id,
                backup_id=backup_id,
            ),
            args_summary={"workbook_id": workbook_id, "backup_id": backup_id},
            default_workbook_id=workbook_id,
        )
