from __future__ import annotations

from typing import Literal

from pydantic import Field, field_validator

from .common import ClientRequestMixin, StrictModel

A1_CELL_PATTERN = r"^\$?[A-Za-z]{1,3}\$?\d{1,7}$"
A1_RANGE_PATTERN = r"^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$"
COLUMN_PATTERN = r"^[A-Za-z]{1,3}$"

ScalarValue = str | int | float | bool | None


class RangeReadValuesRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    value_mode: Literal["raw", "display"] = "raw"
    include_formulas: bool = False
    row_offset: int = Field(default=0, ge=0)
    row_limit: int = Field(default=200, ge=1, le=1000)


class RangeReadValuesData(StrictModel):
    sheet_name: str
    source_range: str
    page_range: str
    total_rows: int
    returned_rows: int
    column_count: int
    values: list[list[ScalarValue]]
    formulas: list[list[str | None]] | None = None
    has_more: bool
    next_row_offset: int | None


class RangeWriteValuesRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str | None = None
    start_cell: str = Field(pattern=A1_CELL_PATTERN)
    values: list[list[ScalarValue]]

    @field_validator("values")
    @classmethod
    def validate_values(cls, values: list[list[ScalarValue]]) -> list[list[ScalarValue]]:
        if not values:
            raise ValueError("values cannot be empty")
        width = len(values[0])
        if width == 0:
            raise ValueError("values rows must be non-empty")
        for row in values:
            if len(row) != width:
                raise ValueError("values must be rectangular")
        return values


class RangeWriteValuesData(StrictModel):
    sheet_name: str
    affected_range: str
    rows_written: int
    columns_written: int
    cells_written: int
    snapshot_id: str


class RangeClearContentsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    scope: str = Field(default="contents")


class RangeClearContentsData(StrictModel):
    sheet_name: str
    affected_range: str
    cells_cleared: int
    scope: str = "contents"
    rollback_supported: bool = True
    warnings: list[str] = []
    snapshot_id: str | None


class RangeCopyRangeRequest(ClientRequestMixin):
    workbook_id: str
    source_sheet: str = Field(max_length=31)
    source_range: str = Field(pattern=A1_RANGE_PATTERN)
    target_sheet: str = Field(max_length=31)
    target_start_cell: str = Field(pattern=A1_CELL_PATTERN)
    paste_mode: Literal["values", "formulas"] = "values"
    target_workbook_id: str | None = None


class RangeCopyRangeData(StrictModel):
    source_sheet: str
    source_range: str
    target_sheet: str
    target_range: str
    rows_copied: int
    columns_copied: int
    cells_copied: int
    paste_mode: Literal["values", "formulas"]
    snapshot_id: str


class RangeInsertRowsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    row_number: int = Field(ge=1, le=1_048_576)
    count: int = Field(default=1, ge=1, le=1000)


class RangeInsertRowsData(StrictModel):
    sheet_name: str
    inserted_at_row: int
    rows_inserted: int
    inserted_range: str
    backup_id: str
    invalidated_snapshots: int


class RangeDeleteRowsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    start_row: int = Field(ge=1, le=1_048_576)
    count: int = Field(default=1, ge=1, le=1000)


class RangeDeleteRowsData(StrictModel):
    sheet_name: str
    deleted_from_row: int
    rows_deleted: int
    deleted_range: str
    backup_id: str
    invalidated_snapshots: int


class RangeInsertColumnsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    column: str = Field(pattern=COLUMN_PATTERN)
    count: int = Field(default=1, ge=1, le=100)


class RangeInsertColumnsData(StrictModel):
    sheet_name: str
    inserted_at_column: str
    columns_inserted: int
    inserted_range: str
    backup_id: str
    invalidated_snapshots: int


class RangeDeleteColumnsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    start_column: str = Field(pattern=COLUMN_PATTERN)
    count: int = Field(default=1, ge=1, le=100)


class RangeDeleteColumnsData(StrictModel):
    sheet_name: str
    deleted_from_column: str
    columns_deleted: int
    deleted_range: str
    backup_id: str
    invalidated_snapshots: int


class RangeSortDataRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    sort_fields: list[dict]
    has_header: bool = False
    case_sensitive: bool = False


class RangeSortDataData(StrictModel):
    sheet_name: str
    sorted_range: str
    rows_sorted: int
    backup_id: str
    snapshot_id: str


class RangeMergeCellsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    across: bool = False


class RangeMergeCellsData(StrictModel):
    sheet_name: str
    merged_range: str
    cells_merged: int
    backup_id: str
    snapshot_id: str


class RangeUnmergeCellsRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)


class RangeUnmergeCellsData(StrictModel):
    sheet_name: str
    affected_range: str
    merge_areas_unmerged: int
    merge_ranges: list[str]
    cells_affected: int
    snapshot_id: str


class RangeManageMergeRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    range: str = Field(pattern=A1_RANGE_PATTERN)
    action: str = Field(pattern="^(merge|unmerge)$")
    across: bool = False


class RangeManageMergeData(StrictModel):
    action: str
    sheet_name: str
    affected_range: str
    cells_affected: int
    merge_areas: int | None = None
    snapshot_id: str | None = None
    backup_id: str | None = None
