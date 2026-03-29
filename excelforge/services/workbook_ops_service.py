from __future__ import annotations

import logging
import os
import time
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.models.workbook_ops_models import (
    LinkInfo,
    SheetExportCsvData,
    SheetExportCsvRequest,
    WorkbookCalculateData,
    WorkbookCalculateRequest,
    WorkbookExportPdfData,
    WorkbookExportPdfRequest,
    WorkbookListLinksData,
    WorkbookListLinksRequest,
    WorkbookRefreshAllData,
    WorkbookRefreshAllRequest,
    WorkbookSaveAsData,
    WorkbookSaveAsRequest,
)
from excelforge.runtime.excel_worker import ExcelWorker

logger = logging.getLogger(__name__)


class WorkbookOpsService:
    def __init__(self, config: AppConfig, worker: ExcelWorker) -> None:
        self._config = config
        self._worker = worker

    def save_as(self, params: WorkbookSaveAsRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj

            try:
                if params.password:
                    workbook.SaveAs(
                        Filename=params.save_as_path,
                        Password=params.password,
                        FileFormat=self._get_file_format(params.file_format),
                    )
                else:
                    workbook.SaveAs(
                        Filename=params.save_as_path,
                        FileFormat=self._get_file_format(params.file_format),
                    )

                size_bytes = None
                if os.path.exists(params.save_as_path):
                    size_bytes = os.path.getsize(params.save_as_path)

                return {
                    "file_path": params.save_as_path,
                    "file_format": params.file_format or "xlsx",
                    "size_bytes": size_bytes,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to save workbook as: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=60, requires_excel=True)

    def refresh_all(self, params: WorkbookRefreshAllRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            start_time = time.time()

            try:
                queries_refreshed = 0
                pivots_refreshed = 0

                try:
                    for qt in workbook.Queries:
                        qt.Refresh()
                        queries_refreshed += 1
                except Exception:
                    pass

                try:
                    for ws in workbook.Worksheets:
                        for pf in ws.PivotTables:
                            pf.RefreshTable()
                            pivots_refreshed += 1
                except Exception:
                    pass

                duration_ms = int((time.time() - start_time) * 1000)

                return {
                    "queries_refreshed": queries_refreshed,
                    "pivots_refreshed": pivots_refreshed,
                    "duration_ms": duration_ms,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to refresh: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=300, requires_excel=True)

    def calculate(self, params: WorkbookCalculateRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            start_time = time.time()

            try:
                sheets_calculated = 0
                for ws in workbook.Worksheets:
                    ws.Calculate()
                    sheets_calculated += 1

                duration_ms = int((time.time() - start_time) * 1000)

                return {
                    "sheets_calculated": sheets_calculated,
                    "duration_ms": duration_ms,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to calculate: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=300, requires_excel=True)

    def list_links(self, params: WorkbookListLinksRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            links: list[dict[str, Any]] = []

            try:
                link_sources = workbook.LinkSources()
                if link_sources:
                    for link in link_sources:
                        try:
                            link_type = self._get_link_type(str(link))
                            links.append({
                                "link_type": link_type,
                                "target": str(link),
                                "update_mode": "manual",
                            })
                        except Exception:
                            continue
            except Exception as exc:
                logger.warning("Failed to list links: %s", exc)

            return {
                "link_count": len(links),
                "links": links,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def export_pdf(self, params: WorkbookExportPdfRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj

            try:
                wb_state = workbook.Windows(1).WindowState
                if not params.include_hidden_sheets:
                    workbook.Windows(1).WindowState = -4137

                workbook.ExportAsFixedFormat(
                    Type=1,
                    Filename=params.file_path,
                    Quality=1,
                    IncludeDocProperties=False,
                    IgnorePrintAreas=False,
                )

                if not params.include_hidden_sheets:
                    workbook.Windows(1).WindowState = wb_state

                size_bytes = None
                if os.path.exists(params.file_path):
                    size_bytes = os.path.getsize(params.file_path)

                return {
                    "file_path": params.file_path,
                    "page_count": None,
                    "size_bytes": size_bytes,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to export PDF: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=120, requires_excel=True)

    def export_csv(self, params: SheetExportCsvRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj

            try:
                ws = workbook.Worksheets(params.sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {params.sheet_name}") from exc

            try:
                used_range = ws.UsedRange
                if not used_range:
                    raise ExcelForgeError(ErrorCode.E400_INVALID_REQUEST, "Sheet has no data")

                csv_content = self._range_to_csv(used_range, params.delimiter, params.include_header)

                with open(params.file_path, "w", encoding="utf-8-sig") as f:
                    f.write(csv_content)

                try:
                    rows_exported = used_range.Rows.Count
                except Exception:
                    rows_exported = 0
                size_bytes = os.path.getsize(params.file_path)

                return {
                    "file_path": params.file_path,
                    "rows_exported": rows_exported,
                    "size_bytes": size_bytes,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to export CSV: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=120, requires_excel=True)

    def _get_file_format(self, file_format: str | None) -> int:
        format_map = {
            "xlsx": 51,
            "xlsm": 52,
            "xlsb": 50,
            "xls": 43,
            "csv": 6,
            "pdf": 0,
        }
        return format_map.get(file_format or "xlsx", 51)

    def _get_link_type(self, link: str) -> str:
        if link.startswith("http"):
            return "web"
        elif link.startswith("file://"):
            return "file"
        elif "[Excel]" in link or ".xls" in link:
            return "excel"
        elif ".dq" in link or ".odc" in link:
            return "data_query"
        else:
            return "other"

    def _range_to_csv(self, range_obj: Any, delimiter: str, include_header: bool) -> str:
        data = range_obj.Value
        if data is None:
            return ""
        rows = []
        for row in data:
            quoted_cells = []
            for cell in row:
                cell_str = str(cell) if cell is not None else ""
                if delimiter in cell_str or '"' in cell_str:
                    cell_str = '"' + cell_str.replace('"', '""') + '"'
                quoted_cells.append(cell_str)
            rows.append(delimiter.join(quoted_cells))
        return "\n".join(rows)
