from __future__ import annotations

from typing import Any

from excelforge.models.analysis_models import (
    AnalysisExportReportRequest,
    AnalysisScanFormulasRequest,
    AnalysisScanHiddenRequest,
    AnalysisScanLinksRequest,
    AnalysisScanStructureRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class AnalysisApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def scan_structure(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AnalysisScanStructureRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="analysis.scan_structure",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.analysis_service.scan_structure(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )

    def scan_formulas(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AnalysisScanFormulasRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name"),
            scan_range=params.get("scan_range"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="analysis.scan_formulas",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.analysis_service.scan_formulas(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "scan_range": req.scan_range,
            },
            default_workbook_id=req.workbook_id,
        )

    def scan_links(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AnalysisScanLinksRequest(
            workbook_id=params.get("workbook_id", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="analysis.scan_links",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.analysis_service.scan_links(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def scan_hidden(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AnalysisScanHiddenRequest(
            workbook_id=params.get("workbook_id", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="analysis.scan_hidden",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.analysis_service.scan_hidden(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def export_report(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = AnalysisExportReportRequest(
            workbook_id=params.get("workbook_id", ""),
            report_format=params.get("report_format", "text"),
            include_formulas=params.get("include_formulas", False),
            include_links=params.get("include_links", False),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="analysis.export_report",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.analysis_service.export_report(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "report_format": req.report_format,
                "include_formulas": req.include_formulas,
                "include_links": req.include_links,
            },
            default_workbook_id=req.workbook_id,
        )
