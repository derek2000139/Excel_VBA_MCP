from __future__ import annotations

from pydantic import Field

from .common import ClientRequestMixin, StrictModel


class WorkbookSaveAsRequest(ClientRequestMixin):
    workbook_id: str
    save_as_path: str
    file_format: str | None = None
    password: str | None = None


class WorkbookSaveAsData(StrictModel):
    file_path: str
    file_format: str
    size_bytes: int | None = None


class WorkbookRefreshAllRequest(ClientRequestMixin):
    workbook_id: str


class WorkbookRefreshAllData(StrictModel):
    queries_refreshed: int
    pivots_refreshed: int
    duration_ms: int


class WorkbookCalculateRequest(ClientRequestMixin):
    workbook_id: str


class WorkbookCalculateData(StrictModel):
    sheets_calculated: int
    duration_ms: int


class WorkbookListLinksRequest(ClientRequestMixin):
    workbook_id: str


class LinkInfo(StrictModel):
    link_type: str
    target: str
    update_mode: str


class WorkbookListLinksData(StrictModel):
    link_count: int
    links: list[LinkInfo]


class WorkbookExportPdfRequest(ClientRequestMixin):
    workbook_id: str
    file_path: str
    include_hidden_sheets: bool = False


class WorkbookExportPdfData(StrictModel):
    file_path: str
    page_count: int | None = None
    size_bytes: int | None = None


class SheetExportCsvRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str
    file_path: str
    delimiter: str = ","
    include_header: bool = True


class SheetExportCsvData(StrictModel):
    file_path: str
    rows_exported: int
    size_bytes: int | None = None
