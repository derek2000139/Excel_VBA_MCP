from __future__ import annotations

import logging
import re
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.models.analysis_models import (
    AnalysisExportReportData,
    AnalysisExportReportRequest,
    AnalysisFormulasData,
    AnalysisHiddenData,
    AnalysisLinksData,
    AnalysisScanFormulasRequest,
    AnalysisScanHiddenRequest,
    AnalysisScanLinksRequest,
    AnalysisScanStructureRequest,
    AnalysisStructureData,
    FormulaCell,
    HiddenInfo,
    LinkInfo,
    NameSummary,
    SheetSummary,
)
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.utils.address_parser import parse_cell

logger = logging.getLogger(__name__)


class AnalysisService:
    def __init__(self, config: AppConfig, worker: ExcelWorker) -> None:
        self._config = config
        self._worker = worker

    def scan_structure(self, params: AnalysisScanStructureRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            sheets: list[SheetSummary] = []
            visible_count = 0
            hidden_count = 0

            target_sheets = (
                [workbook.Worksheets(params.sheet_name)] if params.sheet_name else workbook.Worksheets
            )

            for ws in target_sheets:
                try:
                    is_visible = ws.Visible == -1
                    if is_visible:
                        visible_count += 1
                    else:
                        hidden_count += 1

                    sheets.append(SheetSummary(
                        name=ws.Name,
                        sheet_type=str(ws.Type),
                        is_visible=is_visible,
                        row_count=int(ws.UsedRange.Rows.Count) if ws.UsedRange.Rows.Count > 0 else 0,
                        column_count=int(ws.UsedRange.Columns.Count) if ws.UsedRange.Columns.Count > 0 else 0,
                    ))
                except Exception as exc:
                    logger.warning("Failed to scan sheet %s: %s", ws.Name, exc)
                    continue

            defined_names: list[NameSummary] = []
            try:
                for name in workbook.Names:
                    try:
                        defined_names.append(NameSummary(
                            name=name.Name,
                            scope="workbook",
                            refers_to=str(name.RefersTo),
                        ))
                    except Exception:
                        continue
            except Exception:
                pass

            has_external = False
            try:
                if workbook.LinkSources():
                    has_external = True
            except Exception:
                pass

            return AnalysisStructureData(
                sheet_count=len(sheets),
                visible_sheet_count=visible_count,
                hidden_sheet_count=hidden_count,
                sheets=sheets,
                defined_names=defined_names,
                has_external_links=has_external,
            ).model_dump(mode="json")

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def scan_formulas(self, params: AnalysisScanFormulasRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            formula_cells: list[FormulaCell] = []
            external_count = 0

            target_sheets = (
                [workbook.Worksheets(params.sheet_name)] if params.sheet_name else workbook.Worksheets
            )

            for ws in target_sheets:
                try:
                    if params.scan_range:
                        scan_range = ws.Range(params.scan_range)
                    else:
                        scan_range = ws.UsedRange

                    for cell in scan_range.Cells:
                        try:
                            if cell.HasFormula:
                                formula = str(cell.Formula)
                                formula_cells.append(FormulaCell(
                                    address=cell.Address,
                                    formula=formula,
                                    sheet_name=ws.Name,
                                ))
                                if self._has_external_reference(formula):
                                    external_count += 1
                        except Exception:
                            continue
                except Exception as exc:
                    logger.warning("Failed to scan formulas on sheet %s: %s", ws.Name, exc)
                    continue

            return AnalysisFormulasData(
                total_formulas=len(formula_cells),
                formula_cells=formula_cells[:1000],
                has_external_formulas=external_count > 0,
                external_formula_count=external_count,
            ).model_dump(mode="json")

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def scan_links(self, params: AnalysisScanLinksRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            links: list[LinkInfo] = []

            try:
                link_sources = workbook.LinkSources()
                if link_sources:
                    for link in link_sources:
                        try:
                            link_type = "unknown"
                            if link.startswith("http"):
                                link_type = "web"
                            elif link.startswith("file://"):
                                link_type = "file"
                            elif "Excel" in str(link):
                                link_type = "excel"

                            links.append(LinkInfo(
                                link_type=link_type,
                                address=str(link),
                                target=str(link),
                            ))
                        except Exception:
                            continue
            except Exception as exc:
                logger.warning("Failed to scan links: %s", exc)

            return AnalysisLinksData(
                external_link_count=len(links),
                links=links,
            ).model_dump(mode="json")

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def scan_hidden(self, params: AnalysisScanHiddenRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            hidden_elements: list[HiddenInfo] = []
            hidden_sheet_count = 0
            hidden_names_count = 0
            hidden_rows_total = 0
            hidden_columns_total = 0

            for ws in workbook.Worksheets:
                try:
                    if ws.Visible != -1:
                        hidden_sheet_count += 1
                        hidden_elements.append(HiddenInfo(
                            element_type="sheet",
                            name=ws.Name,
                        ))
                except Exception:
                    continue

            try:
                for name in workbook.Names:
                    try:
                        if name.Visible == False:
                            hidden_names_count += 1
                            hidden_elements.append(HiddenInfo(
                                element_type="name",
                                name=name.Name,
                                scope="workbook",
                            ))
                    except Exception:
                        continue
            except Exception:
                pass

            return AnalysisHiddenData(
                hidden_sheet_count=hidden_sheet_count,
                hidden_names_count=hidden_names_count,
                hidden_rows=hidden_rows_total,
                hidden_columns=hidden_columns_total,
                hidden_elements=hidden_elements,
            ).model_dump(mode="json")

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def export_report(self, params: AnalysisExportReportRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            lines: list[str] = []

            lines.append(f"=== Workbook Analysis Report ===")
            lines.append(f"Name: {workbook.Name}")
            lines.append(f"Sheet Count: {workbook.Worksheets.Count}")

            lines.append(f"\n--- Sheets ---")
            for ws in workbook.Worksheets:
                lines.append(f"  - {ws.Name} (Visible: {ws.Visible == -1})")

            if params.include_formulas:
                lines.append(f"\n--- Formulas (sample) ---")
                for ws in workbook.Worksheets:
                    try:
                        for cell in ws.UsedRange.Cells:
                            if cell.HasFormula:
                                lines.append(f"  {ws.Name}!{cell.Address}: {cell.Formula}")
                    except Exception:
                        continue

            if params.include_links:
                lines.append(f"\n--- External Links ---")
                try:
                    links = workbook.LinkSources()
                    if links:
                        for link in links:
                            lines.append(f"  - {link}")
                    else:
                        lines.append("  No external links found.")
                except Exception:
                    lines.append("  Unable to retrieve links.")

            report_text = "\n".join(lines)

            return AnalysisExportReportData(
                report=report_text,
                format=params.report_format,
            ).model_dump(mode="json")

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    @staticmethod
    def _has_external_reference(formula: str) -> bool:
        patterns = [
            r"\[.*\]",
            r"!",
            r"http[s]?://",
            r"file://",
        ]
        for pattern in patterns:
            if re.search(pattern, formula):
                return True
        return False
