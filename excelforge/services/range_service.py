from __future__ import annotations

import logging
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.services.backup_service import BackupService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.address_parser import (
    CellRef,
    RangeRef,
    column_to_index,
    index_to_column,
    parse_cell,
    parse_range,
    range_to_a1,
    shifted_row_page,
)
from excelforge.utils.value_codec import ensure_rectangular, to_excel_matrix, to_scalar

MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384

# Excel COM 常量
XL_SHIFT_DOWN = -4121
XL_SHIFT_UP = -4162
XL_SHIFT_TO_RIGHT = -4161
XL_SHIFT_TO_LEFT = -4159
XL_SORT_ASCENDING = 1
XL_SORT_DESCENDING = 2
XL_YES = 1
XL_NO = 2

logger = logging.getLogger(__name__)


class RangeService:
    def __init__(
        self,
        config: AppConfig,
        worker: ExcelWorker,
        snapshot_service: SnapshotService,
        backup_service: BackupService,
    ) -> None:
        self._config = config
        self._worker = worker
        self._snapshot_service = snapshot_service
        self._backup_service = backup_service

    def read_values(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        value_mode: str,
        include_formulas: bool,
        row_offset: int,
        row_limit: int,
    ) -> dict[str, Any]:
        source_ref = parse_range(range_address)
        if source_ref.cell_count > self._config.limits.max_read_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Read range too large: {source_ref.cell_count}")
        if row_limit > self._config.limits.max_read_rows:
            raise ExcelForgeError(
                ErrorCode.E413_RANGE_TOO_LARGE,
                f"row_limit exceeds max_read_rows={self._config.limits.max_read_rows}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            _, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            page_ref = shifted_row_page(source_ref, row_offset, row_limit)
            returned_rows = max(0, min(row_limit, source_ref.rows - row_offset))
            if returned_rows <= 0:
                return {
                    "sheet_name": sheet_name,
                    "source_range": range_address,
                    "page_range": range_to_a1(page_ref),
                    "total_rows": source_ref.rows,
                    "returned_rows": 0,
                    "column_count": source_ref.cols,
                    "values": [],
                    "formulas": [] if include_formulas else None,
                    "has_more": False,
                    "next_row_offset": None,
                }

            if value_mode != "display" and not include_formulas:
                rng = ws.Range(range_to_a1(page_ref))
                raw = rng.Value2
                if isinstance(raw, tuple):
                    values = [[to_scalar(c) for c in row] for row in raw]
                else:
                    values = [[to_scalar(raw)]]
            else:
                values = []
                formulas = []
                for row in range(page_ref.start.row, page_ref.end.row + 1):
                    row_values: list[Any] = []
                    row_formulas: list[str | None] = []
                    for col in range(page_ref.start.col, page_ref.end.col + 1):
                        cell = ws.Cells(row, col)
                        row_values.append(str(cell.Text) if value_mode == "display" else to_scalar(cell.Value2))
                        if include_formulas:
                            row_formulas.append(str(cell.Formula) if bool(cell.HasFormula) else None)
                    values.append(row_values)
                    if include_formulas:
                        formulas.append(row_formulas)

            has_more = (row_offset + returned_rows) < source_ref.rows
            next_offset = (row_offset + returned_rows) if has_more else None
            return {
                "sheet_name": sheet_name,
                "source_range": range_address,
                "page_range": range_to_a1(page_ref),
                "total_rows": source_ref.rows,
                "returned_rows": returned_rows,
                "column_count": source_ref.cols,
                "values": values,
                "formulas": formulas if include_formulas else None,
                "has_more": has_more,
                "next_row_offset": next_offset,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def write_values(
        self,
        *,
        workbook_id: str,
        sheet_name: str | None,
        start_cell: str,
        values: list[list[Any]],
    ) -> dict[str, Any]:
        try:
            rows, cols = ensure_rectangular(values)
        except ExcelForgeError as exc:
            if exc.code == ErrorCode.E400_INVALID_ARGUMENT:
                raise ExcelForgeError(ErrorCode.E400_EMPTY_VALUES, "values cannot be empty and must be rectangular") from exc
            raise

        cell_count = rows * cols
        if cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Write range too large: {cell_count}")

        start = parse_cell(start_cell)
        target_ref = RangeRef(start=start, end=CellRef(row=start.row + rows - 1, col=start.col + cols - 1))
        target_range = range_to_a1(target_ref)

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=target_range,
                source_tool="range.write_values",
            )
            try:
                ws.Range(target_range).Value = to_excel_matrix(values)
            except Exception as exc:
                self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to write range values: {exc}") from exc

            return {
                "sheet_name": sheet_name,
                "affected_range": target_range,
                "rows_written": rows,
                "columns_written": cols,
                "cells_written": cell_count,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def clear_contents(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        scope: str = "contents",
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Clear range too large: {target_ref.cell_count}")

        rollback_supported = scope in ("contents", "all")
        snapshot_id = None

        def op(ctx: Any) -> dict[str, Any]:
            nonlocal snapshot_id, rollback_supported
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)

            if rollback_supported:
                snapshot_id = self._snapshot_service.create_snapshot(
                    workbook=handle,
                    worksheet=ws,
                    range_address=range_address,
                    source_tool="range.clear_contents",
                )

            warnings: list[str] = []
            try:
                rng = ws.Range(range_address)
                if scope == "contents":
                    rng.ClearContents()
                elif scope == "formats":
                    rng.ClearFormats()
                    rollback_supported = False
                    warnings.append("此操作不可回滚")
                else:
                    rng.Clear()
                    warnings.append("快照仅能恢复值和公式，格式变更不可回滚")
            except Exception as exc:
                if snapshot_id:
                    self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to clear range: {exc}") from exc

            return {
                "sheet_name": sheet_name,
                "affected_range": range_address,
                "cells_cleared": target_ref.cell_count,
                "scope": scope,
                "rollback_supported": rollback_supported,
                "warnings": warnings,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def copy_range(
        self,
        *,
        workbook_id: str,
        source_sheet: str,
        source_range: str,
        target_sheet: str,
        target_start_cell: str,
        paste_mode: str = "values",
        target_workbook_id: str | None = None,
    ) -> dict[str, Any]:
        source_ref = parse_range(source_range)
        if source_ref.cell_count > self._config.limits.max_copy_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Copy range too large: {source_ref.cell_count}")

        target_start = parse_cell(target_start_cell)
        target_ref = RangeRef(
            start=target_start,
            end=CellRef(row=target_start.row + source_ref.rows - 1, col=target_start.col + source_ref.cols - 1),
        )
        target_range = range_to_a1(target_ref)

        target_wb_id = target_workbook_id or workbook_id

        def op(ctx: Any) -> dict[str, Any]:
            source_handle, source_ws = self._require_sheet(ctx, workbook_id, source_sheet)

            if target_wb_id != workbook_id:
                target_handle = ctx.registry.get(target_wb_id)
                if target_handle is None:
                    raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Target workbook not open: {target_wb_id}")
            else:
                target_handle = source_handle

            _, target_ws = self._require_sheet(ctx, target_wb_id, target_sheet)
            self._ensure_write_allowed(target_handle, target_ws)

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=target_handle,
                worksheet=target_ws,
                range_address=target_range,
                source_tool="range.copy_range",
            )
            try:
                src_rng = source_ws.Range(source_range)
                tgt_rng = target_ws.Range(target_range)
                if paste_mode == "formulas":
                    tgt_rng.Formula = src_rng.Formula
                else:
                    tgt_rng.Value = src_rng.Value
            except Exception as exc:
                self._restore_snapshot_best_effort(ctx, snapshot_id, target_wb_id)
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to copy range: {exc}") from exc

            return {
                "source_sheet": source_sheet,
                "source_range": source_range,
                "target_sheet": target_sheet,
                "target_range": target_range,
                "rows_copied": source_ref.rows,
                "columns_copied": source_ref.cols,
                "cells_copied": source_ref.cell_count,
                "paste_mode": paste_mode,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def insert_rows(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        row_number: int,
        count: int,
    ) -> dict[str, Any]:
        if row_number < 1 or row_number > MAX_EXCEL_ROWS:
            raise ExcelForgeError(ErrorCode.E400_ROW_OUT_OF_RANGE, f"row_number out of range: {row_number}")
        if count > self._config.limits.max_insert_rows:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                f"count exceeds max_insert_rows={self._config.limits.max_insert_rows}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.insert_rows",
                description=f"插入行 {row_number}:{row_number + count - 1}",
            )
            insert_range = ws.Rows(f"{row_number}:{row_number + count - 1}")
            try:
                insert_range.Insert(Shift=XL_SHIFT_DOWN)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to insert rows: {exc}") from exc
            invalidated = self._snapshot_service.expire_sheet_snapshots(workbook_id, sheet_name)
            return {
                "sheet_name": sheet_name,
                "inserted_at_row": row_number,
                "rows_inserted": count,
                "inserted_range": f"{row_number}:{row_number + count - 1}",
                "backup_id": backup_id,
                "invalidated_snapshots": invalidated,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def delete_rows(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        start_row: int,
        count: int,
    ) -> dict[str, Any]:
        end_row = start_row + count - 1
        if start_row < 1 or start_row > MAX_EXCEL_ROWS or end_row > MAX_EXCEL_ROWS:
            raise ExcelForgeError(
                ErrorCode.E400_ROW_OUT_OF_RANGE,
                f"Row range out of range: {start_row}:{end_row}",
            )
        if count > self._config.limits.max_insert_rows:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                f"count exceeds max_insert_rows={self._config.limits.max_insert_rows}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.delete_rows",
                description=f"删除行 {start_row}:{end_row}",
            )
            delete_range = ws.Rows(f"{start_row}:{end_row}")
            try:
                delete_range.Delete(Shift=XL_SHIFT_UP)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to delete rows: {exc}") from exc
            invalidated = self._snapshot_service.expire_sheet_snapshots(workbook_id, sheet_name)
            return {
                "sheet_name": sheet_name,
                "deleted_from_row": start_row,
                "rows_deleted": count,
                "deleted_range": f"{start_row}:{end_row}",
                "backup_id": backup_id,
                "invalidated_snapshots": invalidated,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def insert_columns(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        column: str,
        count: int,
    ) -> dict[str, Any]:
        start_idx = self._parse_column_index(column)
        end_idx = start_idx + count - 1
        if end_idx > MAX_EXCEL_COLUMNS:
            raise ExcelForgeError(
                ErrorCode.E400_COLUMN_OUT_OF_RANGE,
                f"Column range out of range: {column}:{index_to_column(end_idx)}",
            )
        if count > self._config.limits.max_insert_columns:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                f"count exceeds max_insert_columns={self._config.limits.max_insert_columns}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            start_col = index_to_column(start_idx)
            end_col = index_to_column(end_idx)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.insert_columns",
                description=f"插入列 {start_col}:{end_col}",
            )
            try:
                ws.Columns(f"{start_col}:{end_col}").Insert(Shift=XL_SHIFT_TO_RIGHT)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to insert columns: {exc}") from exc
            invalidated = self._snapshot_service.expire_sheet_snapshots(workbook_id, sheet_name)
            return {
                "sheet_name": sheet_name,
                "inserted_at_column": start_col,
                "columns_inserted": count,
                "inserted_range": f"{start_col}:{end_col}",
                "backup_id": backup_id,
                "invalidated_snapshots": invalidated,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def delete_columns(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        start_column: str,
        count: int,
    ) -> dict[str, Any]:
        start_idx = self._parse_column_index(start_column)
        end_idx = start_idx + count - 1
        if end_idx > MAX_EXCEL_COLUMNS:
            raise ExcelForgeError(
                ErrorCode.E400_COLUMN_OUT_OF_RANGE,
                f"Column range out of range: {start_column}:{index_to_column(end_idx)}",
            )
        if count > self._config.limits.max_insert_columns:
            raise ExcelForgeError(
                ErrorCode.E400_INVALID_ARGUMENT,
                f"count exceeds max_insert_columns={self._config.limits.max_insert_columns}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            start_col = index_to_column(start_idx)
            end_col = index_to_column(end_idx)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.delete_columns",
                description=f"删除列 {start_col}:{end_col}",
            )
            try:
                ws.Columns(f"{start_col}:{end_col}").Delete(Shift=XL_SHIFT_TO_LEFT)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to delete columns: {exc}") from exc
            invalidated = self._snapshot_service.expire_sheet_snapshots(workbook_id, sheet_name)
            return {
                "sheet_name": sheet_name,
                "deleted_from_column": start_col,
                "columns_deleted": count,
                "deleted_range": f"{start_col}:{end_col}",
                "backup_id": backup_id,
                "invalidated_snapshots": invalidated,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def sort_data(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        sort_fields: list[dict],
        has_header: bool,
        case_sensitive: bool,
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Sort range too large: {target_ref.cell_count}")

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.sort_data",
                description=f"排序 {range_address}",
            )
            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=range_address,
                source_tool="range.sort_data",
            )
            try:
                sort_range = ws.Range(range_address)
                ws.Sort.SortFields.Clear()

                for field in sort_fields:
                    col_letter = field["column"].upper()
                    col_idx = self._column_letter_to_index(col_letter)
                    key_range = ws.Range(
                        f"{col_letter}{target_ref.start.row}:{col_letter}{target_ref.end.row}"
                    )
                    order = XL_SORT_DESCENDING if field.get("descending", False) else XL_SORT_ASCENDING
                    ws.Sort.SortFields.Add(Key=key_range, Order=order)

                ws.Sort.SetRange(sort_range)
                ws.Sort.Header = XL_YES if has_header else XL_NO
                ws.Sort.MatchCase = case_sensitive
                ws.Sort.Apply()
            except Exception as exc:
                self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to sort range: {exc}") from exc

            return {
                "sheet_name": sheet_name,
                "sorted_range": range_address,
                "rows_sorted": target_ref.rows,
                "backup_id": backup_id,
                "snapshot_id": snapshot_id,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def merge_cells(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        across: bool,
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="range.merge_cells",
                description=f"合并单元格 {range_address}",
            )
            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=range_address,
                source_tool="range.merge_cells",
            )
            try:
                ws.Range(range_address).Merge(Across=across)
            except Exception as exc:
                self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to merge cells: {exc}") from exc

            return {
                "sheet_name": sheet_name,
                "merged_range": range_address,
                "cells_merged": target_ref.cell_count,
                "backup_id": backup_id,
                "snapshot_id": snapshot_id,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def unmerge_cells(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Unmerge range too large: {target_ref.cell_count}")

        def op(ctx: Any) -> dict[str, Any]:
            handle, ws = self._require_sheet(ctx, workbook_id, sheet_name)
            self._ensure_write_allowed(handle, ws)
            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=range_address,
                source_tool="range.unmerge_cells",
            )

            merge_areas: dict[str, Any] = {}
            cells_affected = 0
            try:
                rng = ws.Range(range_address)
                for cell in rng.Cells:
                    try:
                        if cell.MergeCells:
                            area = cell.MergeArea
                            addr = area.Address.replace("$", "")
                            if addr not in merge_areas:
                                merge_areas[addr] = area
                                cells_affected += int(area.Count)
                    except Exception:
                        continue

                for area in merge_areas.values():
                    try:
                        area.UnMerge()
                    except Exception:
                        pass
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to unmerge cells: {exc}") from exc

            return {
                "sheet_name": sheet_name,
                "affected_range": range_address,
                "merge_areas_unmerged": len(merge_areas),
                "merge_ranges": list(merge_areas.keys()),
                "cells_affected": cells_affected,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def manage_merge(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        action: str,
        across: bool = False,
    ) -> dict[str, Any]:
        if action == "merge":
            result = self.merge_cells(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                range_address=range_address,
                across=across,
            )
            return {
                "action": "merge",
                "sheet_name": result["sheet_name"],
                "affected_range": result["merged_range"],
                "cells_affected": result["cells_merged"],
                "merge_areas": 1,
                "snapshot_id": result["snapshot_id"],
                "backup_id": result["backup_id"],
            }
        elif action == "unmerge":
            result = self.unmerge_cells(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                range_address=range_address,
            )
            return {
                "action": "unmerge",
                "sheet_name": result["sheet_name"],
                "affected_range": result["affected_range"],
                "cells_affected": result["cells_affected"],
                "merge_areas": result["merge_areas_unmerged"],
                "snapshot_id": result["snapshot_id"],
                "backup_id": None,
            }
        else:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid action: {action}")

    @staticmethod
    def _column_letter_to_index(column: str) -> int:
        result = 0
        for c in column.upper():
            result = result * 26 + (ord(c) - ord("A") + 1)
        return result

    @staticmethod
    def _parse_column_index(column: str) -> int:
        try:
            idx = column_to_index(column.upper())
        except ExcelForgeError as exc:
            raise ExcelForgeError(ErrorCode.E400_COLUMN_OUT_OF_RANGE, f"Invalid column: {column}") from exc
        if idx < 1 or idx > MAX_EXCEL_COLUMNS:
            raise ExcelForgeError(ErrorCode.E400_COLUMN_OUT_OF_RANGE, f"Invalid column: {column}")
        return idx

    def _require_sheet(self, ctx: Any, workbook_id: str, sheet_name: str | None) -> tuple[WorkbookHandle, Any]:
        handle = ctx.registry.get(workbook_id)
        if handle is None:
            if ctx.registry.is_stale_workbook_id(workbook_id):
                raise ExcelForgeError(
                    ErrorCode.E410_WORKBOOK_STALE,
                    "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                )
            raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
        workbook = handle.workbook_obj
        if sheet_name is None or sheet_name == "None" or sheet_name == "":
            ws = workbook.ActiveSheet
            if ws is None:
                ws = workbook.Worksheets(1)
        else:
            try:
                ws = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc
        return handle, ws

    def _ensure_write_allowed(self, handle: WorkbookHandle, ws: Any) -> None:
        workbook = handle.workbook_obj
        if bool(workbook.ReadOnly):
            raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")
        if bool(ws.ProtectContents):
            raise ExcelForgeError(ErrorCode.E403_SHEET_PROTECTED, "Worksheet is protected")

    def _restore_snapshot_best_effort(self, ctx: Any, snapshot_id: str, workbook_id: str) -> None:
        try:
            meta, payload = self._snapshot_service.load_snapshot(snapshot_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                logger.warning("Cannot restore snapshot %s: workbook %s not found", snapshot_id, workbook_id)
                return
            workbook = handle.workbook_obj
            ws = workbook.Worksheets(str(meta["sheet_name"]))
            self._snapshot_service.restore_snapshot(workbook=handle, worksheet=ws, snapshot_payload=payload)
        except Exception as exc:
            logger.warning("Best-effort snapshot restore failed for snapshot_id=%s: %s", snapshot_id, exc)
            return
