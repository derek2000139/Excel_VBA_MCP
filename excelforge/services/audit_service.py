from __future__ import annotations

import getpass
import socket
from dataclasses import dataclass
from typing import Any

from excelforge.config import AppConfig
from excelforge.persistence.audit_repo import AuditRecord, AuditRepository


@dataclass
class AuditContext:
    tool_name: str
    operation_id: str
    workbook_id: str | None
    file_path: str | None
    started_at: str
    duration_ms: int
    success: bool
    code: str
    message: str
    affected_sheet: str | None = None
    affected_range: str | None = None
    snapshot_id: str | None = None
    args_summary: dict[str, object] | None = None
    client_request_id: str | None = None
    client_name: str | None = None
    actor_id: str | None = None


class AuditService:
    def __init__(self, config: AppConfig, repo: AuditRepository) -> None:
        self._config = config
        self._repo = repo

    def record_operation(self, entry: AuditContext) -> None:
        actor_id = entry.actor_id or self._config.server.actor_id
        record = AuditRecord(
            operation_id=entry.operation_id,
            tool_name=entry.tool_name,
            workbook_id=entry.workbook_id,
            file_path=entry.file_path,
            actor_id=actor_id,
            os_user=getpass.getuser(),
            machine_name=socket.gethostname(),
            client_name=entry.client_name,
            client_request_id=entry.client_request_id,
            started_at=entry.started_at,
            duration_ms=entry.duration_ms,
            success=entry.success,
            code=entry.code,
            message=entry.message,
            affected_sheet=entry.affected_sheet,
            affected_range=entry.affected_range,
            snapshot_id=entry.snapshot_id,
            args_summary=entry.args_summary,
        )
        self._repo.insert(record)

    def list_operations(
        self,
        *,
        workbook_id: str | None,
        tool_name: str | None,
        success_only: bool,
        limit: int,
        offset: int,
        operation_id: str | None = None,
    ) -> dict[str, Any]:
        if operation_id:
            op = self._repo.get_operation(operation_id)
            if op is None:
                return {
                    "total": 0,
                    "has_more": False,
                    "next_offset": None,
                    "items": [],
                }
            return {
                "total": 1,
                "has_more": False,
                "next_offset": None,
                "items": [op],
            }

        total, items = self._repo.list_operations(
            workbook_id=workbook_id,
            tool_name=tool_name,
            success_only=success_only,
            limit=limit,
            offset=offset,
        )
        has_more = (offset + len(items)) < total
        next_offset = (offset + len(items)) if has_more else None
        return {
            "total": total,
            "has_more": has_more,
            "next_offset": next_offset,
            "items": items,
        }
