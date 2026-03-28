from __future__ import annotations

from typing import Any

from excelforge.runtime_api.context import RuntimeApiContext


class FormatApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def set_style(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        range_address = str(params.get("range", ""))
        style = params.get("style") or {}
        return self._ctx.run_operation(
            method_name="format.set_style",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.format_service.set_range_style(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                range_address=range_address,
                style=style,
            ),
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range_address,
            },
            default_workbook_id=workbook_id,
        )

    def auto_fit(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        columns = params.get("range")
        return self._ctx.run_operation(
            method_name="format.auto_fit",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.format_service.auto_fit_columns(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                columns=columns,
            ),
            args_summary={"workbook_id": workbook_id, "sheet_name": sheet_name, "range": columns},
            default_workbook_id=workbook_id,
        )
