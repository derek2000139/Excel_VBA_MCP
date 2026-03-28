from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .common import ClientRequestMixin, StrictModel
from .range_models import ScalarValue


class SheetInspectStructureRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    sample_rows: int = Field(default=5, ge=1, le=20)
    scan_rows: int = Field(default=10, ge=1, le=30)
    max_profile_columns: int = Field(default=50, ge=1, le=200)


class SheetHeaderProfile(StrictModel):
    column_index: int
    column_letter: str
    header_text: str
    inferred_type: Literal["empty", "text", "number", "date", "boolean", "mixed"]
    sample_values: list[ScalarValue]
    non_empty_count: int
    unique_count: int
    number_format: str | None
    has_formulas: bool


class SheetInspectStructureData(StrictModel):
    sheet_name: str
    used_range: str
    total_rows: int
    total_columns: int
    header_row_candidate: int
    heuristic: bool
    profile_truncated: bool
    headers: list[SheetHeaderProfile]
    sample_data: list[list[ScalarValue]]
    has_auto_filter: bool
    frozen_panes: bool
    merged_cells_count: int
    protected: bool


class SheetCreateRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    position: Literal["first", "last"] = "last"


class SheetCreateData(StrictModel):
    workbook_id: str
    sheet_name: str
    sheet_index: int
    total_sheets: int


class SheetRenameRequest(ClientRequestMixin):
    workbook_id: str
    current_name: str = Field(max_length=31)
    new_name: str = Field(max_length=31)


class SheetRenameData(StrictModel):
    workbook_id: str
    previous_name: str
    new_name: str
    sheet_index: int


class SheetDeleteSheetRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    preview: bool = False
    confirm_token: str = ""

    @model_validator(mode="after")
    def validate_confirm_token(self):
        if not self.preview and not self.confirm_token:
            raise ValueError("confirm_token is required when preview is false")
        return self


class SheetDeleteCrossReferences(StrictModel):
    referenced_by_count: int
    referencing_sheets: list[str]


class SheetPreviewDeleteData(StrictModel):
    workbook_id: str
    sheet_name: str
    used_range: str
    total_rows: int
    total_columns: int
    is_last_visible_sheet: bool
    cross_references: SheetDeleteCrossReferences
    active_snapshots_count: int
    can_delete: bool
    block_reason: str | None
    confirm_token: str | None
    confirm_token_expires_at: str | None
    warnings: list[str]


class SheetDeleteSheetData(StrictModel):
    workbook_id: str
    deleted_sheet: str | None = None
    backup_id: str | None = None
    invalidated_snapshots: int | None = None
    remaining_sheets: list[str] | None = None
    preview: bool = False
    preview_data: SheetPreviewDeleteData | None = None


class SheetSetAutoFilterRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    action: str
    range: str | None = None
    filters: list[dict] | None = None


class SheetSetAutoFilterAppliedFilter(StrictModel):
    column: str
    operator: str
    value: str | int | float | None = None
    value2: str | int | float | None = None


class SheetSetAutoFilterData(StrictModel):
    sheet_name: str
    action: str
    normalized_action: str
    filter_range: str | None
    auto_filter_active: bool
    applied_filters: list[SheetSetAutoFilterAppliedFilter]


class SheetGetRulesRequest(ClientRequestMixin):
    workbook_id: str
    sheet_name: str = Field(max_length=31)
    rule_type: str = "conditional_formats"
    range: str = ""
    limit: int = Field(default=100, ge=1, le=500)


class SheetConditionalFormatItem(StrictModel):
    applies_to: str
    type: str
    operator: str | None
    formula1: str | None
    formula2: str | None
    priority: int | None
    stop_if_true: bool | None


class SheetDataValidationItem(StrictModel):
    type: str
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    show_input_message: bool
    prompt_title: str | None
    prompt: str | None
