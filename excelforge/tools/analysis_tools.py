from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.analysis_models import (
    AnalysisExportReportRequest,
    AnalysisScanFormulasRequest,
    AnalysisScanHiddenRequest,
    AnalysisScanLinksRequest,
    AnalysisScanStructureRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_analysis_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="analysis.scan_structure")
    def analysis_scan_structure(
        workbook_id: str,
        sheet_name: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = AnalysisScanStructureRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="analysis.scan_structure",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.analysis_service.scan_structure(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("analysis.scan_structure", "analysis_tools", "analysis")

    @mcp.tool(name="analysis.scan_formulas")
    def analysis_scan_formulas(
        workbook_id: str,
        sheet_name: str | None = None,
        scan_range: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = AnalysisScanFormulasRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            scan_range=scan_range,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="analysis.scan_formulas",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.analysis_service.scan_formulas(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "scan_range": req.scan_range,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("analysis.scan_formulas", "analysis_tools", "analysis")

    @mcp.tool(name="analysis.scan_links")
    def analysis_scan_links(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = AnalysisScanLinksRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="analysis.scan_links",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.analysis_service.scan_links(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("analysis.scan_links", "analysis_tools", "analysis")

    @mcp.tool(name="analysis.scan_hidden")
    def analysis_scan_hidden(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = AnalysisScanHiddenRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="analysis.scan_hidden",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.analysis_service.scan_hidden(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("analysis.scan_hidden", "analysis_tools", "analysis")

    @mcp.tool(name="analysis.export_report")
    def analysis_export_report(
        workbook_id: str,
        report_format: str = "text",
        include_formulas: bool = False,
        include_links: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = AnalysisExportReportRequest(
            workbook_id=workbook_id,
            report_format=report_format,
            include_formulas=include_formulas,
            include_links=include_links,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="analysis.export_report",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.analysis_service.export_report(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "report_format": req.report_format,
                "include_formulas": req.include_formulas,
                "include_links": req.include_links,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("analysis.export_report", "analysis_tools", "analysis")
