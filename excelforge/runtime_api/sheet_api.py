from __future__ import annotations

from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.models.sheet_models import SheetDeleteSheetRequest, SheetInspectStructureRequest, SheetSetAutoFilterRequest
from excelforge.runtime_api.context import RuntimeApiContext


class SheetApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def inspect(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = SheetInspectStructureRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            sample_rows=int(params.get("sample_rows", 5)),
            scan_rows=int(params.get("scan_rows", 10)),
            max_profile_columns=int(params.get("max_profile_columns", 50)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="sheet.inspect",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.sheet_service.inspect_structure(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                sample_rows=req.sample_rows,
                scan_rows=req.scan_rows,
                max_profile_columns=req.max_profile_columns,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "sample_rows": req.sample_rows,
                "scan_rows": req.scan_rows,
                "max_profile_columns": req.max_profile_columns,
            },
            default_workbook_id=req.workbook_id,
        )

    def create(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        position = str(params.get("position", "last")).lower()
        reference_sheet = str(params.get("reference_sheet", "")).strip() or None
        client_request_id = params.get("client_request_id")

        if position in {"first", "last"}:
            return self._ctx.run_operation(
                method_name="sheet.create",
                actor_id=actor_id,
                client_request_id=client_request_id,
                operation_fn=lambda: self._ctx.services.sheet_service.create_sheet(
                    workbook_id=workbook_id,
                    sheet_name=sheet_name,
                    position=position,
                ),
                args_summary={
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "position": position,
                    "reference_sheet": reference_sheet,
                },
                default_workbook_id=workbook_id,
            )

        if position not in {"before", "after"} or not reference_sheet:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                "position must be first/last, or before/after with reference_sheet",
            )

        def op() -> dict[str, Any]:
            def worker_op(worker_ctx: Any) -> dict[str, Any]:
                handle = worker_ctx.registry.get(workbook_id)
                if handle is None:
                    raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
                workbook = handle.workbook_obj
                try:
                    ref_ws = workbook.Worksheets(reference_sheet)
                except Exception as exc:
                    raise ExcelForgeError(
                        ErrorCode.E404_SHEET_NOT_FOUND,
                        f"Sheet not found: {reference_sheet}",
                    ) from exc

                if position == "before":
                    ws = workbook.Sheets.Add(Before=ref_ws)
                else:
                    ws = workbook.Sheets.Add(After=ref_ws)
                ws.Name = sheet_name
                return {
                    "workbook_id": workbook_id,
                    "sheet_name": str(ws.Name),
                    "sheet_index": int(ws.Index),
                    "total_sheets": int(workbook.Sheets.Count),
                }

            return self._ctx.services.worker.submit(
                worker_op,
                timeout_seconds=self._ctx.services.config.limits.operation_timeout_seconds,
                requires_excel=True,
            )

        return self._ctx.run_operation(
            method_name="sheet.create",
            actor_id=actor_id,
            client_request_id=client_request_id,
            operation_fn=op,
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "position": position,
                "reference_sheet": reference_sheet,
            },
            default_workbook_id=workbook_id,
        )

    def rename(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        new_name = str(params.get("new_name", ""))
        return self._ctx.run_operation(
            method_name="sheet.rename",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.sheet_service.rename_sheet(
                workbook_id=workbook_id,
                current_name=sheet_name,
                new_name=new_name,
            ),
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "new_name": new_name,
            },
            default_workbook_id=workbook_id,
        )

    def preview_delete(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        return self._ctx.run_operation(
            method_name="sheet.preview_delete",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.sheet_service.delete_sheet(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                preview=True,
                confirm_token="",
            ),
            args_summary={"workbook_id": workbook_id, "sheet_name": sheet_name},
            default_workbook_id=workbook_id,
        )

    def delete(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = SheetDeleteSheetRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            preview=False,
            confirm_token=params.get("confirm_token", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="sheet.delete",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.sheet_service.delete_sheet(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                preview=False,
                confirm_token=req.confirm_token,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def auto_filter(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = SheetSetAutoFilterRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            action=params.get("action", ""),
            range=params.get("range"),
            filters=params.get("filters"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="sheet.auto_filter",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.sheet_service.set_auto_filter(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                action=req.action,
                range_address=req.range,
                filters=req.filters,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "action": req.action,
                "range": req.range,
            },
            default_workbook_id=req.workbook_id,
        )

    def get_conditional_formats(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        range_address = str(params.get("range", "") or "")
        limit = int(params.get("limit", 100))
        return self._ctx.run_operation(
            method_name="sheet.get_conditional_formats",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.sheet_service.get_rules(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                rule_type="conditional_formats",
                range_address=range_address,
                limit=limit,
            ),
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range_address,
                "limit": limit,
            },
            default_workbook_id=workbook_id,
        )

    def get_data_validations(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        range_address = str(params.get("range", "") or "")
        limit = int(params.get("limit", 100))
        return self._ctx.run_operation(
            method_name="sheet.get_data_validations",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.sheet_service.get_rules(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                rule_type="data_validations",
                range_address=range_address,
                limit=limit,
            ),
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range_address,
                "limit": limit,
            },
            default_workbook_id=workbook_id,
        )
