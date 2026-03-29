from __future__ import annotations

from typing import Any

from excelforge.models.workbook_ops_models import (
    SheetExportCsvRequest,
    WorkbookCalculateRequest,
    WorkbookExportPdfRequest,
    WorkbookListLinksRequest,
    WorkbookRefreshAllRequest,
    WorkbookSaveAsRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class WorkbookOpsApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def save_as(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookSaveAsRequest(
            workbook_id=params.get("workbook_id", ""),
            save_as_path=params.get("save_as_path", ""),
            file_format=params.get("file_format"),
            password=params.get("password"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.save_as",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.save_as(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "save_as_path": req.save_as_path,
                "file_format": req.file_format,
            },
            default_workbook_id=req.workbook_id,
        )

    def refresh_all(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookRefreshAllRequest(
            workbook_id=params.get("workbook_id", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.refresh_all",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.refresh_all(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def calculate(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookCalculateRequest(
            workbook_id=params.get("workbook_id", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.calculate",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.calculate(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def list_links(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookListLinksRequest(
            workbook_id=params.get("workbook_id", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.list_links",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.list_links(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def export_pdf(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookExportPdfRequest(
            workbook_id=params.get("workbook_id", ""),
            file_path=params.get("file_path", ""),
            include_hidden_sheets=params.get("include_hidden_sheets", False),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.export_pdf",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.export_pdf(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "file_path": req.file_path,
                "include_hidden_sheets": req.include_hidden_sheets,
            },
            default_workbook_id=req.workbook_id,
        )

    def export_csv(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = SheetExportCsvRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            file_path=params.get("file_path", ""),
            delimiter=params.get("delimiter", ","),
            include_header=params.get("include_header", True),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="sheet.export_csv",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_ops_service.export_csv(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "file_path": req.file_path,
                "delimiter": req.delimiter,
                "include_header": req.include_header,
            },
            default_workbook_id=req.workbook_id,
        )