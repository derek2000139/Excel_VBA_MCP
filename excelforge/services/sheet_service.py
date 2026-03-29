from __future__ import annotations

from dataclasses import dataclass
from datetime import timedelta
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.backup_service import BackupService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.address_parser import CellRef, RangeRef, index_to_column, range_to_a1
from excelforge.utils.ids import generate_id
from excelforge.utils.timestamps import utc_now
from excelforge.utils.value_codec import to_scalar

INVALID_SHEET_NAME_CHARS = set("\\/?*[]")


@dataclass
class ConfirmTokenEntry:
    workbook_id: str
    sheet_name: str
    expires_at: Any


class SheetService:
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
        self._confirm_tokens: dict[str, ConfirmTokenEntry] = {}

    def inspect_structure(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        sample_rows: int,
        scan_rows: int,
        max_profile_columns: int,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            workbook = handle.workbook_obj
            try:
                ws = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            used = ws.UsedRange
            start_row = int(used.Row)
            start_col = int(used.Column)
            total_rows = int(used.Rows.Count)
            total_cols = int(used.Columns.Count)

            if total_rows <= 0 or total_cols <= 0:
                return {
                    "sheet_name": sheet_name,
                    "used_range": "A1",
                    "total_rows": 0,
                    "total_columns": 0,
                    "header_row_candidate": 1,
                    "heuristic": True,
                    "profile_truncated": False,
                    "headers": [],
                    "sample_data": [],
                    "has_auto_filter": bool(ws.AutoFilterMode),
                    "frozen_panes": self._is_frozen_panes(workbook),
                    "merged_cells_count": 0,
                    "protected": bool(ws.ProtectContents),
                }

            profile_cols = min(total_cols, max_profile_columns)
            profile_truncated = profile_cols < total_cols
            scan_limit = min(scan_rows, total_rows)
            data_scan_end = min(start_row + total_rows - 1, start_row + scan_limit - 1)

            header_row_candidate = start_row
            best_score = -1
            for row in range(start_row, data_scan_end + 1):
                non_empty = 0
                text_like = 0
                for col in range(start_col, start_col + profile_cols):
                    value = ws.Cells(row, col).Value2
                    if value not in (None, ""):
                        non_empty += 1
                        if isinstance(value, str):
                            text_like += 1
                score = non_empty * 2 + text_like
                if score > best_score:
                    best_score = score
                    header_row_candidate = row

            headers: list[dict[str, Any]] = []
            sample_data: list[list[Any]] = []
            data_start_row = header_row_candidate + 1
            data_end_row = min(start_row + total_rows - 1, data_start_row + scan_rows - 1)

            for row in range(data_start_row, min(data_start_row + sample_rows, data_end_row + 1)):
                sample_data.append(
                    [
                        to_scalar(ws.Cells(row, col).Value2)
                        for col in range(start_col, start_col + profile_cols)
                    ]
                )

            for offset in range(profile_cols):
                col = start_col + offset
                header_cell = ws.Cells(header_row_candidate, col).Value2
                header_text = "" if header_cell in (None, "") else str(header_cell)

                sample_values: list[Any] = []
                typed_values: list[Any] = []
                has_formulas = False
                unique_values: set[str] = set()
                non_empty_count = 0
                first_number_format: str | None = None

                for row in range(data_start_row, data_end_row + 1):
                    cell = ws.Cells(row, col)
                    scalar = to_scalar(cell.Value2)
                    if bool(cell.HasFormula):
                        has_formulas = True
                    if scalar not in (None, ""):
                        non_empty_count += 1
                        typed_values.append(scalar)
                        unique_values.add(str(scalar))
                        if len(sample_values) < sample_rows:
                            sample_values.append(scalar)
                        if first_number_format is None:
                            try:
                                first_number_format = str(cell.NumberFormat)
                            except Exception:
                                first_number_format = None

                inferred = infer_type(typed_values, first_number_format)
                headers.append(
                    {
                        "column_index": offset + 1,
                        "column_letter": index_to_column(col),
                        "header_text": header_text,
                        "inferred_type": inferred,
                        "sample_values": sample_values,
                        "non_empty_count": non_empty_count,
                        "unique_count": len(unique_values),
                        "number_format": first_number_format,
                        "has_formulas": has_formulas,
                    }
                )

            used_range = range_to_a1(
                RangeRef(
                    start=CellRef(start_row, start_col),
                    end=CellRef(start_row + total_rows - 1, start_col + total_cols - 1),
                )
            )

            return {
                "sheet_name": sheet_name,
                "used_range": used_range,
                "total_rows": total_rows,
                "total_columns": total_cols,
                "header_row_candidate": header_row_candidate,
                "heuristic": True,
                "profile_truncated": profile_truncated,
                "headers": headers,
                "sample_data": sample_data,
                "has_auto_filter": bool(ws.AutoFilterMode),
                "frozen_panes": self._is_frozen_panes(workbook),
                "merged_cells_count": self._count_merged_cells(used),
                "protected": bool(ws.ProtectContents),
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def create_sheet(self, *, workbook_id: str, sheet_name: str, position: str) -> dict[str, Any]:
        self._validate_sheet_name(sheet_name)

        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            wb = handle.workbook_obj
            self._ensure_workbook_writable(wb)
            if int(wb.Sheets.Count) >= self._config.limits.max_create_sheets:
                raise ExcelForgeError(
                    ErrorCode.E423_FEATURE_NOT_SUPPORTED,
                    f"Sheet count exceeds max_create_sheets={self._config.limits.max_create_sheets}",
                )
            if self._sheet_exists(wb, sheet_name):
                raise ExcelForgeError(ErrorCode.E409_SHEET_NAME_EXISTS, f"Sheet already exists: {sheet_name}")
            if position == "first":
                ws = wb.Sheets.Add(Before:=wb.Sheets(1))
            elif position == "last":
                ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            else:
                # after_active: Excel 原生行为，在当前激活表之后插入
                ws = wb.Sheets.Add(After:=wb.ActiveSheet)
            ws.Name = sheet_name
            return {
                "workbook_id": workbook_id,
                "sheet_name": str(ws.Name),
                "sheet_index": int(ws.Index),
                "total_sheets": int(wb.Sheets.Count),
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def rename_sheet(self, *, workbook_id: str, current_name: str, new_name: str) -> dict[str, Any]:
        self._validate_sheet_name(new_name)

        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            wb = handle.workbook_obj
            self._ensure_workbook_writable(wb)
            try:
                ws = wb.Worksheets(current_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {current_name}") from exc

            if current_name.lower() != new_name.lower() and self._sheet_exists(wb, new_name):
                raise ExcelForgeError(ErrorCode.E409_SHEET_NAME_EXISTS, f"Sheet already exists: {new_name}")

            previous_name = str(ws.Name)
            ws.Name = new_name
            self._snapshot_service.rename_sheet_snapshot_refs(workbook_id, previous_name, new_name)
            return {
                "workbook_id": workbook_id,
                "previous_name": previous_name,
                "new_name": str(ws.Name),
                "sheet_index": int(ws.Index),
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def delete_sheet(self, *, workbook_id: str, sheet_name: str, preview: bool = False, confirm_token: str = "") -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            wb = handle.workbook_obj
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            if preview:
                used = ws.UsedRange
                start_row = int(used.Row)
                start_col = int(used.Column)
                total_rows = int(used.Rows.Count)
                total_cols = int(used.Columns.Count)
                used_range = range_to_a1(
                    RangeRef(
                        start=CellRef(start_row, start_col),
                        end=CellRef(start_row + total_rows - 1, start_col + total_cols - 1),
                    )
                )

                ref_count, referencing_sheets = self._estimate_cross_references(wb, sheet_name)
                active_snapshots = self._snapshot_service.count_active_for_sheet(workbook_id, sheet_name)
                is_last_visible = self._is_last_visible_sheet(wb, ws)
                can_delete = (not is_last_visible) and (not bool(wb.ReadOnly))
                block_reason: str | None = None
                if bool(wb.ReadOnly):
                    block_reason = "Workbook is read-only"
                elif is_last_visible:
                    block_reason = "Cannot delete the last visible worksheet"

                confirm_token_out: str | None = None
                confirm_expires_at: str | None = None
                if can_delete:
                    confirm_token_out, confirm_expires_at = self._issue_confirm_token(workbook_id, sheet_name)

                warnings: list[str] = []
                if ref_count > 0:
                    warnings.append("其他工作表存在对此表的公式引用，删除后相关公式可能变为 #REF!")
                if active_snapshots > 0:
                    warnings.append("该表上的活跃快照将在删除后失效。")

                return {
                    "workbook_id": workbook_id,
                    "sheet_name": sheet_name,
                    "used_range": used_range,
                    "total_rows": total_rows,
                    "total_columns": total_cols,
                    "is_last_visible_sheet": is_last_visible,
                    "cross_references": {
                        "referenced_by_count": ref_count,
                        "referencing_sheets": referencing_sheets,
                    },
                    "active_snapshots_count": active_snapshots,
                    "can_delete": can_delete,
                    "block_reason": block_reason,
                    "confirm_token": confirm_token_out,
                    "confirm_token_expires_at": confirm_expires_at,
                    "warnings": warnings,
                    "preview": True,
                }

            self._consume_confirm_token(confirm_token, workbook_id, sheet_name)
            self._ensure_workbook_writable(wb)

            if self._is_last_visible_sheet(wb, ws):
                raise ExcelForgeError(
                    ErrorCode.E409_CANNOT_DELETE_LAST_SHEET,
                    "Cannot delete the last visible worksheet",
                )

            backup_id, backup_warnings = self._backup_service.create_backup(
                workbook=handle,
                source_tool="sheet.delete_sheet",
                description=f"删除工作表 {sheet_name}",
            )
            try:
                ws.Delete()
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to delete sheet: {exc}") from exc

            invalidated = self._snapshot_service.expire_sheet_snapshots(workbook_id, sheet_name)
            remaining = [str(item.Name) for item in wb.Worksheets]
            return {
                "workbook_id": workbook_id,
                "deleted_sheet": sheet_name,
                "backup_id": backup_id,
                "invalidated_snapshots": invalidated,
                "remaining_sheets": remaining,
                "preview": False,
                "__warnings__": backup_warnings,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def _issue_confirm_token(self, workbook_id: str, sheet_name: str) -> tuple[str, str]:
        self._cleanup_expired_tokens()
        token = generate_id("ctok")
        expires_at_dt = utc_now() + timedelta(minutes=self._config.backup.confirm_token_ttl_minutes)
        self._confirm_tokens[token] = ConfirmTokenEntry(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            expires_at=expires_at_dt,
        )
        return token, expires_at_dt.isoformat().replace("+00:00", "Z")

    def _consume_confirm_token(self, token: str, workbook_id: str, sheet_name: str) -> None:
        self._cleanup_expired_tokens()
        entry = self._confirm_tokens.pop(token, None)
        if entry is None:
            raise ExcelForgeError(ErrorCode.E409_CONFIRM_TOKEN_INVALID, "confirm_token is invalid or expired")
        if entry.workbook_id != workbook_id or entry.sheet_name != sheet_name:
            raise ExcelForgeError(ErrorCode.E409_CONFIRM_TOKEN_INVALID, "confirm_token does not match target sheet")
        if entry.expires_at <= utc_now():
            raise ExcelForgeError(ErrorCode.E409_CONFIRM_TOKEN_INVALID, "confirm_token is expired")

    def _cleanup_expired_tokens(self) -> None:
        now = utc_now()
        expired = [token for token, entry in self._confirm_tokens.items() if entry.expires_at <= now]
        for token in expired:
            self._confirm_tokens.pop(token, None)

    def set_auto_filter(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        action: str,
        range_address: str | None = None,
        filters: list[dict] | None = None,
    ) -> dict[str, Any]:
        ACTION_ALIAS_MAP = {
            "enable": "enable",
            "add": "enable",
            "on": "enable",
            "disable": "disable",
            "remove": "disable",
            "off": "disable",
        }
        normalized_action = ACTION_ALIAS_MAP.get(action.lower(), action)

        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            wb = handle.workbook_obj
            self._ensure_workbook_writable(wb)
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            if normalized_action == "enable":
                if range_address:
                    ws.Range(range_address).AutoFilter()
                else:
                    used = ws.UsedRange
                    ws.Range(used.Address).AutoFilter()

                if filters:
                    # 使用 Range.AutoFilter 直接设置筛选条件
                    # 只支持单个筛选条件，通过指定列和条件值
                    f = filters[0]
                    col_idx = self._column_letter_to_index(f["column"])
                    op_val = self._operator_to_excel(f["operator"], f.get("value"), f.get("value2"))
                    
                    # 获取目标列的范围（整列）
                    col_letter = self._column_index_to_letter(col_idx)
                    col_range = ws.Range(f"{col_letter}:{col_letter}")
                    
                    # 对该列应用筛选
                    col_range.AutoFilter(Field=1, Criteria1=op_val[1])
            else:
                ws.AutoFilterMode = False

            af = ws.AutoFilter
            applied: list[dict] = []
            try:
                if af is not None and af.Filters.Count > 0:
                    for i in range(1, af.Filters.Count + 1):
                        f = af.Filters(i)
                        applied.append({
                            "column": self._column_index_to_letter(i),
                            "operator": str(f.Operator),
                            "value": f.Criteria1,
                            "value2": f.Criteria2,
                        })
            except Exception:
                pass

            return {
                "sheet_name": sheet_name,
                "action": action,
                "normalized_action": normalized_action,
                "filter_range": range_address,
                "auto_filter_active": bool(ws.AutoFilterMode),
                "applied_filters": applied,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def get_rules(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        rule_type: str = "conditional_formats",
        range_address: str = "",
        limit: int = 100,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = self._require_workbook(ctx, workbook_id)
            wb = handle.workbook_obj
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            items: list[dict] = []
            count = 0

            if rule_type == "data_validations":
                try:
                    dvs = ws.DataValidations
                except Exception:
                    dvs = []
                try:
                    for dv in dvs:
                        if count >= limit:
                            break
                        items.append({
                            "type": str(dv.Type),
                            "formula1": dv.Formula1 if hasattr(dv, "Formula1") else None,
                            "formula2": dv.Formula2 if hasattr(dv, "Formula2") else None,
                            "allow_blank": bool(dv.AllowBlank) if hasattr(dv, "AllowBlank") else False,
                            "show_input_message": bool(dv.ShowInputMessage) if hasattr(dv, "ShowInputMessage") else False,
                            "prompt_title": dv.PromptTitle if hasattr(dv, "PromptTitle") else None,
                            "prompt": dv.Prompt if hasattr(dv, "Prompt") else None,
                        })
                        count += 1
                except Exception:
                    pass
            else:
                try:
                    cf = ws.ConditionalFormats
                except Exception:
                    cf = []
                if range_address:
                    try:
                        target = ws.Range(range_address)
                        for c in target.ConditionalFormats:
                            if count >= limit:
                                break
                            items.append(self._cf_to_dict(c))
                            count += 1
                    except Exception:
                        pass
                else:
                    try:
                        for c in cf:
                            if count >= limit:
                                break
                            items.append(self._cf_to_dict(c))
                            count += 1
                    except Exception:
                        pass

            return {
                "sheet_name": sheet_name,
                "rule_type": rule_type,
                "total_rules": count,
                "truncated": count >= limit,
                "items": items,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    @staticmethod
    def _column_letter_to_index(col: str) -> int:
        result = 0
        for c in col.upper():
            result = result * 26 + (ord(c) - ord("A") + 1)
        return result

    @staticmethod
    def _column_index_to_letter(idx: int) -> str:
        result = ""
        while idx > 0:
            idx, remainder = divmod(idx - 1, 26)
            result = chr(65 + remainder) + result
        return result

    @staticmethod
    def _operator_to_excel(op: str, value: Any, value2: Any) -> tuple:
        OPERATOR_MAP = {
            "equals": (1, f"={value}", None, False),
            "not_equals": (2, f"={value}", None, False),
            "contains": (1, f"=*{value}*", None, False),
            "begins_with": (1, f"={value}*", None, False),
            "ends_with": (1, f"=*{value}", None, False),
            "greater_than": (1, f">{value}", None, False),
            "less_than": (1, f"<{value}", None, False),
            "between": (1, f">={value}", f"<={value}", True),
            "blank": (1, "=", None, False),
            "non_blank": (1, "<>", None, False),
        }
        return OPERATOR_MAP.get(op, (1, f"={value}", None, False))

    @staticmethod
    def _cf_to_dict(cf: Any) -> dict:
        try:
            return {
                "applies_to": str(cf.AppliesTo.Address) if hasattr(cf.AppliesTo, "Address") else None,
                "type": str(cf.Type) if hasattr(cf, "Type") else None,
                "operator": str(cf.Operator) if hasattr(cf, "Operator") else None,
                "formula1": str(cf.Formula1) if hasattr(cf, "Formula1") and cf.Formula1 else None,
                "formula2": str(cf.Formula2) if hasattr(cf, "Formula2") and cf.Formula2 else None,
                "priority": int(cf.Priority) if hasattr(cf, "Priority") else None,
                "stop_if_true": bool(cf.StopIfTrue) if hasattr(cf, "StopIfTrue") else None,
            }
        except Exception:
            return {"applies_to": None, "type": None, "operator": None, "formula1": None, "formula2": None, "priority": None, "stop_if_true": None}

    @staticmethod
    def _is_last_visible_sheet(workbook: Any, target_ws: Any) -> bool:
        visible_sheets = [ws for ws in workbook.Worksheets if int(ws.Visible) != 0]
        return int(target_ws.Visible) != 0 and len(visible_sheets) <= 1

    @staticmethod
    def _estimate_cross_references(workbook: Any, target_sheet_name: str) -> tuple[int, list[str]]:
        patterns = [f"{target_sheet_name}!", f"'{target_sheet_name}'!"]
        referencing_sheets: list[str] = []
        ref_count = 0

        for ws in workbook.Worksheets:
            if str(ws.Name) == target_sheet_name:
                continue
            used = ws.UsedRange
            if used is None:
                continue
            used_rows = int(used.Rows.Count)
            if used_rows <= 0:
                continue
            scan_rows = min(used_rows, 200)
            scan_range = ws.Range(
                ws.Cells(int(used.Row), int(used.Column)),
                ws.Cells(int(used.Row) + scan_rows - 1, int(used.Column) + int(used.Columns.Count) - 1),
            )
            formulas = scan_range.Formula
            found = count_formula_mentions(formulas, patterns)
            if found > 0:
                ref_count += found
                referencing_sheets.append(str(ws.Name))

        return ref_count, referencing_sheets

    @staticmethod
    def _is_frozen_panes(workbook: Any) -> bool:
        try:
            return bool(workbook.Windows(1).FreezePanes)
        except Exception:
            return False

    @staticmethod
    def _count_merged_cells(used_range: Any) -> int:
        try:
            merge_areas = used_range.MergeAreas
            total = 0
            for area in merge_areas:
                total += int(area.Count)
            return total
        except Exception:
            return 0

    @staticmethod
    def _validate_sheet_name(name: str) -> None:
        if not name or len(name) > 31:
            raise ExcelForgeError(ErrorCode.E400_INVALID_SHEET_NAME, f"Invalid sheet name: {name}")
        if any(ch in INVALID_SHEET_NAME_CHARS for ch in name):
            raise ExcelForgeError(ErrorCode.E400_INVALID_SHEET_NAME, f"Invalid sheet name: {name}")

    def _require_workbook(self, ctx: Any, workbook_id: str) -> Any:
        handle = ctx.registry.get(workbook_id)
        if handle is None:
            if ctx.registry.is_stale_workbook_id(workbook_id):
                raise ExcelForgeError(
                    ErrorCode.E410_WORKBOOK_STALE,
                    "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                )
            raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
        return handle

    @staticmethod
    def _sheet_exists(workbook: Any, sheet_name: str) -> bool:
        for ws in workbook.Worksheets:
            if str(ws.Name).lower() == sheet_name.lower():
                return True
        return False

    @staticmethod
    def _ensure_workbook_writable(workbook: Any) -> None:
        if bool(workbook.ReadOnly):
            raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")


def infer_type(values: list[Any], number_format: str | None) -> str:
    if not values:
        return "empty"

    type_set: set[str] = set()
    for value in values:
        if isinstance(value, bool):
            type_set.add("boolean")
        elif isinstance(value, (int, float)):
            type_set.add("number")
        elif isinstance(value, str):
            type_set.add("text")
        else:
            type_set.add("mixed")

    if len(type_set) == 1:
        inferred = next(iter(type_set))
    else:
        inferred = "mixed"
    if inferred == "number" and number_format:
        nf = number_format.lower()
        if "yy" in nf or "dd" in nf or "mm" in nf:
            return "date"
    return inferred


def count_formula_mentions(formulas: Any, patterns: list[str]) -> int:
    if formulas is None:
        return 0
    if isinstance(formulas, str):
        return 1 if any(p in formulas for p in patterns) else 0
    if isinstance(formulas, (tuple, list)):
        total = 0
        for item in formulas:
            total += count_formula_mentions(item, patterns)
        return total
    return 0
