from __future__ import annotations

import re
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.models.format_models import StyleModel
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.utils.address_parser import column_to_index, index_to_column, parse_range

HEX_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")
COLUMN_SPAN_RE = re.compile(r"^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$")

HORIZONTAL_ALIGNMENT_MAP: dict[str, int] = {
    "left": -4131,
    "center": -4108,
    "right": -4152,
    "general": 1,
}
VERTICAL_ALIGNMENT_MAP: dict[str, int] = {
    "top": -4160,
    "center": -4108,
    "middle": -4108,
    "bottom": -4107,
}
BORDER_WEIGHT_MAP: dict[str, int] = {
    "thin": 1,
    "medium": -4138,
    "thick": 4,
}
OUTLINE_BORDERS = (7, 8, 9, 10)
ALL_BORDERS = (7, 8, 9, 10, 11, 12)


class FormatService:
    def __init__(self, config: AppConfig, worker: ExcelWorker) -> None:
        self._config = config
        self._worker = worker

    def set_range_style(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        style: dict[str, Any] | None | StyleModel,
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(
                ErrorCode.E413_RANGE_TOO_LARGE,
                f"Format range too large: {target_ref.cell_count}",
            )
        if style is None:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                "At least one style property must be provided",
            )
        if isinstance(style, StyleModel):
            style_dict = style.model_dump(exclude_none=True)
        else:
            style_dict = style

        def op(ctx: Any) -> dict[str, Any]:
            workbook, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(workbook, ws)
            rng = ws.Range(range_address)

            properties_set: list[str] = []

            if style_dict.get("font_name") or style_dict.get("font_size") or style_dict.get("font_bold") or style_dict.get("font_italic") or style_dict.get("font_color"):
                font_obj = rng.Font
                if style_dict.get("font_bold") is not None:
                    font_obj.Bold = bool(style_dict["font_bold"])
                if style_dict.get("font_italic") is not None:
                    font_obj.Italic = bool(style_dict["font_italic"])
                if style_dict.get("font_size") is not None:
                    font_obj.Size = float(style_dict["font_size"])
                if style_dict.get("font_name"):
                    font_obj.Name = str(style_dict["font_name"])
                if style_dict.get("font_color"):
                    font_obj.Color = self._hex_to_excel_color(str(style_dict["font_color"]))
                properties_set.append("font")

            if style_dict.get("fill_color"):
                rng.Interior.Color = self._hex_to_excel_color(str(style_dict["fill_color"]))
                properties_set.append("fill")

            if style_dict.get("horizontal_alignment") or style_dict.get("vertical_alignment") or style_dict.get("wrap_text"):
                if style_dict.get("horizontal_alignment") is not None:
                    rng.HorizontalAlignment = HORIZONTAL_ALIGNMENT_MAP[str(style_dict["horizontal_alignment"])]
                if style_dict.get("vertical_alignment") is not None:
                    rng.VerticalAlignment = VERTICAL_ALIGNMENT_MAP[str(style_dict["vertical_alignment"])]
                if style_dict.get("wrap_text") is not None:
                    rng.WrapText = bool(style_dict["wrap_text"])
                properties_set.append("alignment")

            if style_dict.get("number_format"):
                rng.NumberFormat = style_dict["number_format"]
                properties_set.append("number_format")

            if style_dict.get("border_style"):
                border_style = str(style_dict["border_style"])
                if border_style == "none":
                    rng.Borders.LineStyle = -4142
                    properties_set.append("border")
                elif border_style in BORDER_WEIGHT_MAP:
                    for idx in ALL_BORDERS:
                        border_obj = rng.Borders(idx)
                        border_obj.LineStyle = 1
                        border_obj.Weight = BORDER_WEIGHT_MAP[border_style]
                        border_obj.Color = 0
                    properties_set.append("border")
                else:
                    raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Unsupported border style: {border_style}")

            return {
                "sheet_name": sheet_name,
                "affected_range": range_address,
                "cells_affected": target_ref.cell_count,
                "properties_set": properties_set,
                "__warnings__": [
                    "Format changes are not included in snapshots and cannot be rolled back automatically.",
                ],
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def auto_fit_columns(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        columns: str | None,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            workbook, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(workbook, ws)

            if columns:
                start_col, end_col = self._parse_column_span(columns)
            else:
                used = ws.UsedRange
                used_count = int(used.Columns.Count)
                if used_count <= 0:
                    return {
                        "sheet_name": sheet_name,
                        "columns_adjusted": [],
                        "column_count": 0,
                        "__warnings__": [
                            "Column width changes are not included in snapshots and cannot be rolled back automatically.",
                        ],
                    }
                start_col = int(used.Column)
                end_col = start_col + used_count - 1

            col_expr = f"{index_to_column(start_col)}:{index_to_column(end_col)}"
            ws.Columns(col_expr).AutoFit()
            adjusted = [index_to_column(i) for i in range(start_col, end_col + 1)]
            return {
                "sheet_name": sheet_name,
                "columns_adjusted": adjusted,
                "column_count": len(adjusted),
                "__warnings__": [
                    "Column width changes are not included in snapshots and cannot be rolled back automatically.",
                ],
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    @staticmethod
    def _hex_to_excel_color(hex_color: str) -> int:
        raw = hex_color.lstrip("#")
        if not HEX_COLOR_RE.fullmatch(raw):
            raise ExcelForgeError(ErrorCode.E400_INVALID_COLOR, f"Invalid color value: {hex_color}")
        r = int(raw[0:2], 16)
        g = int(raw[2:4], 16)
        b = int(raw[4:6], 16)
        return r + (g << 8) + (b << 16)

    @staticmethod
    def _parse_column_span(columns: str) -> tuple[int, int]:
        m = COLUMN_SPAN_RE.fullmatch(columns.strip())
        if not m:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid columns span: {columns}")
        start = column_to_index(m.group(1))
        end = column_to_index(m.group(2) or m.group(1))
        if end < start:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid columns span: {columns}")
        return start, end

    @staticmethod
    def _require_sheet(ctx: Any, workbook_id: str, sheet_name: str) -> tuple[Any, Any]:
        handle = ctx.registry.get(workbook_id)
        if handle is None:
            if ctx.registry.is_stale_workbook_id(workbook_id):
                raise ExcelForgeError(
                    ErrorCode.E410_WORKBOOK_STALE,
                    "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                )
            raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
        workbook = handle.workbook_obj
        try:
            ws = workbook.Worksheets(sheet_name)
        except Exception as exc:
            raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc
        return workbook, ws

    @staticmethod
    def _ensure_write_allowed(workbook: Any, ws: Any) -> None:
        if bool(workbook.ReadOnly):
            raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")
        if bool(ws.ProtectContents):
            raise ExcelForgeError(ErrorCode.E403_SHEET_PROTECTED, "Worksheet is protected")

    def manage(
        self,
        *,
        action: str,
        workbook_id: str,
        sheet_name: str,
        range: str = "",
        style: dict[str, Any] | None = None,
        columns: str | None = None,
    ) -> dict[str, Any]:
        if action == "set_style":
            if not range:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "range required for set_style action")
            if style is None:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "style required for set_style action")
            return self.set_range_style(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                range_address=range,
                style=style,
            )
        elif action == "auto_fit_columns":
            return self.auto_fit_columns(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                columns=columns,
            )
        else:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid action: {action}")
