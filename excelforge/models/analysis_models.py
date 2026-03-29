from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class AnalysisScanStructureRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str | None = None


class SheetSummary(StrictModel):
    name: str
    sheet_type: str
    is_visible: bool
    row_count: int
    column_count: int


class NameSummary(StrictModel):
    name: str
    scope: str
    refers_to: str


class AnalysisStructureData(StrictModel):
    sheet_count: int
    visible_sheet_count: int
    hidden_sheet_count: int
    sheets: list[SheetSummary]
    defined_names: list[NameSummary]
    has_external_links: bool


class AnalysisScanFormulasRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str | None = None
    scan_range: str | None = None


class FormulaCell(StrictModel):
    address: str
    formula: str
    sheet_name: str


class AnalysisFormulasData(StrictModel):
    total_formulas: int
    formula_cells: list[FormulaCell]
    has_external_formulas: bool
    external_formula_count: int


class AnalysisScanLinksRequest(ClientRequestMixin):
    workbook_id: str


class LinkInfo(StrictModel):
    link_type: str
    address: str
    target: str


class AnalysisLinksData(StrictModel):
    external_link_count: int
    links: list[LinkInfo]


class AnalysisScanHiddenRequest(ClientRequestMixin):
    workbook_id: str


class HiddenInfo(StrictModel):
    element_type: str
    name: str
    scope: str | None = None


class AnalysisHiddenData(StrictModel):
    hidden_sheet_count: int
    hidden_names_count: int
    hidden_rows: int
    hidden_columns: int
    hidden_elements: list[HiddenInfo]


class AnalysisExportReportRequest(ClientRequestMixin):
    workbook_id: str
    report_format: str = "text"
    include_formulas: bool = False
    include_links: bool = False


class AnalysisExportReportData(StrictModel):
    report: str
    format: str
