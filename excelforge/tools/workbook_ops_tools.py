from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.workbook_ops_models import (
    SheetExportCsvRequest,
    WorkbookCalculateRequest,
    WorkbookExportPdfRequest,
    WorkbookListLinksRequest,
    WorkbookRefreshAllRequest,
    WorkbookSaveAsRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_workbook_ops_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="workbook.save_as")
    def workbook_save_as(
        workbook_id: str,
        save_as_path: str,
        file_format: str | None = None,
        password: str | None = None,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookSaveAsRequest(
            workbook_id=workbook_id,
            save_as_path=save_as_path,
            file_format=file_format,
            password=password,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.save_as",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.save_as(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "save_as_path": req.save_as_path,
                "file_format": req.file_format,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.save_as", "workbook_ops_tools", "workbook_ops")

    @mcp.tool(name="workbook.refresh_all")
    def workbook_refresh_all(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookRefreshAllRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.refresh_all",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.refresh_all(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.refresh_all", "workbook_ops_tools", "workbook_ops")

    @mcp.tool(name="workbook.calculate")
    def workbook_calculate(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookCalculateRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.calculate",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.calculate(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.calculate", "workbook_ops_tools", "workbook_ops")

    @mcp.tool(name="workbook.list_links")
    def workbook_list_links(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookListLinksRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.list_links",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.list_links(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.list_links", "workbook_ops_tools", "workbook_ops")

    @mcp.tool(name="workbook.export_pdf")
    def workbook_export_pdf(
        workbook_id: str,
        file_path: str,
        include_hidden_sheets: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = WorkbookExportPdfRequest(
            workbook_id=workbook_id,
            file_path=file_path,
            include_hidden_sheets=include_hidden_sheets,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="workbook.export_pdf",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.export_pdf(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "file_path": req.file_path,
                "include_hidden_sheets": req.include_hidden_sheets,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("workbook.export_pdf", "workbook_ops_tools", "workbook_ops")

    @mcp.tool(name="sheet.export_csv")
    def sheet_export_csv(
        workbook_id: str,
        sheet_name: str,
        file_path: str,
        delimiter: str = ",",
        include_header: bool = True,
        client_request_id: str = "",
    ) -> dict:
        req = SheetExportCsvRequest(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            file_path=file_path,
            delimiter=delimiter,
            include_header=include_header,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="sheet.export_csv",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.workbook_ops_service.export_csv(params=req),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "file_path": req.file_path,
                "delimiter": req.delimiter,
                "include_header": req.include_header,
            },
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("sheet.export_csv", "workbook_ops_tools", "workbook_ops")