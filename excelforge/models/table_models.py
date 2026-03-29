from __future__ import annotations

from typing import Literal

from pydantic import Field, field_validator

from .common import ClientRequestMixin, StrictModel


class TableListRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str | None = None


class TableColumnInfo(StrictModel):
    name: str
    index: int
    field_type: str


class TableStyleInfo(StrictModel):
    name: str
    id: int


class TableInfo(StrictModel):
    name: str
    sheet_name: str
    range_address: str
    columns: list[TableColumnInfo]
    total_row_count: int
    data_row_count: int
    has_total_row: bool
    style: TableStyleInfo | None = None


class TableListData(StrictModel):
    tables: list[TableInfo]
    total_count: int


class TableCreateRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str
    range_address: str
    table_name: str | None = None
    has_header: bool = True
    style_name: str | None = None


class TableCreateData(StrictModel):
    name: str
    sheet_name: str
    range_address: str
    columns: list[TableColumnInfo]
    total_row_count: int
    data_row_count: int
    has_header: bool
    has_total_row: bool
    style: TableStyleInfo | None = None
    snapshot_id: str | None = None


class TableInspectRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    sheet_name: str | None = None


class TableInspectData(StrictModel):
    name: str
    sheet_name: str
    range_address: str
    columns: list[TableColumnInfo]
    total_row_count: int
    data_row_count: int
    has_header: bool
    has_total_row: bool
    style: TableStyleInfo | None = None


class TableResizeRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    new_range_address: str
    sheet_name: str | None = None


class TableResizeData(StrictModel):
    name: str
    old_range_address: str
    new_range_address: str
    columns: list[TableColumnInfo]
    total_row_count: int
    data_row_count: int
    has_total_row: bool
    snapshot_id: str | None = None


class TableRenameRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    new_name: str
    sheet_name: str | None = None


class TableRenameData(StrictModel):
    old_name: str
    new_name: str
    sheet_name: str
    range_address: str


class TableSetStyleRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    style_name: str | None = None
    sheet_name: str | None = None


class TableSetStyleData(StrictModel):
    name: str
    old_style: TableStyleInfo | None = None
    new_style: TableStyleInfo | None = None
    sheet_name: str


class TableToggleTotalRowRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    sheet_name: str | None = None


class TableToggleTotalRowData(StrictModel):
    name: str
    sheet_name: str
    has_total_row: bool
    snapshot_id: str | None = None


class TableDeleteRequest(ClientRequestMixin):
    workbook_id: str
    table_name: str
    sheet_name: str | None = None


class TableDeleteData(StrictModel):
    name: str
    sheet_name: str
    data_preserved: bool
    data_range_address: str
    snapshot_id: str | None = None
