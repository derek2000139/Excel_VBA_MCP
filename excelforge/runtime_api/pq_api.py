from __future__ import annotations

from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime_api.context import RuntimeApiContext


class PqApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    @staticmethod
    def _not_enabled() -> dict[str, Any]:
        raise ExcelForgeError(
            ErrorCode.E423_FEATURE_NOT_SUPPORTED,
            "Power Query domain is not enabled in this version",
        )

    def list_queries(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        return self._ctx.run_operation(
            method_name="pq.list_queries",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=self._not_enabled,
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )

    def get_query_code(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        query_name = params.get("query_name")
        return self._ctx.run_operation(
            method_name="pq.get_query_code",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=self._not_enabled,
            args_summary={"workbook_id": workbook_id, "query_name": query_name},
            default_workbook_id=workbook_id,
        )

    def update_query(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        query_name = params.get("query_name")
        return self._ctx.run_operation(
            method_name="pq.update_query",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=self._not_enabled,
            args_summary={"workbook_id": workbook_id, "query_name": query_name},
            default_workbook_id=workbook_id,
        )

    def refresh(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        query_name = params.get("query_name")
        return self._ctx.run_operation(
            method_name="pq.refresh",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=self._not_enabled,
            args_summary={"workbook_id": workbook_id, "query_name": query_name},
            default_workbook_id=workbook_id,
        )

    def list_connections(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = params.get("workbook_id")
        return self._ctx.run_operation(
            method_name="pq.list_connections",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=self._not_enabled,
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )
