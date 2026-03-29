from __future__ import annotations

import re
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.address_parser import index_to_column, parse_cell_address, parse_range
from excelforge.utils.calculation_waiter import check_dynamic_array_support, wait_for_calculation
from excelforge.utils.value_codec import to_scalar

A1_REF_RE = re.compile(r"\$?[A-Za-z]{1,3}\$?\d{1,7}")
INVALID_FORMULA_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
ERROR_VALUES = {
    -2146826252: "#SPILL!",
    -2146826281: "#REF!",
    -2146826246: "#NAME?",
    -2146826259: "#DIV/0!",
    -2146826255: "#N/A",
    -2146826273: "#NULL!",
    -2146826256: "#NUM!",
    -2146826257: "#VALUE!",
}


class FormulaService:
    def __init__(self, config: AppConfig, worker: ExcelWorker, snapshot_service: SnapshotService) -> None:
        self._config = config
        self._worker = worker
        self._snapshot_service = snapshot_service

    def validate_expression(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        anchor_cell: str,
        formula: str,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
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
                _ = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            starts_with_equals = formula.startswith("=")
            length_valid = 1 <= len(formula) <= 8192
            no_invalid_chars = INVALID_FORMULA_CHARS.search(formula) is None
            balanced_parens = formula.count("(") == formula.count(")")
            syntax_valid = starts_with_equals and length_valid and no_invalid_chars and balanced_parens

            english_style = not bool(re.search(r"[\u4e00-\u9fff]", formula))
            references = A1_REF_RE.findall(formula)

            warnings: list[str] = []
            if ";" in formula:
                warnings.append("Formula uses semicolon separators; English formula style expects commas.")
            if "FormulaLocal" in formula:
                warnings.append("FormulaLocal is not supported in v0.1.")
            if "{" in formula or "}" in formula:
                warnings.append("Array formula patterns detected; array formulas are not supported in v0.1.")
            if "#" in formula:
                warnings.append("Dynamic array spill references may be incompatible in v0.1.")

            notes = [
                "Validation is basic and not a full Excel parser.",
                f"Anchor cell context: {anchor_cell}",
            ]

            return {
                "syntax_valid": bool(syntax_valid),
                "starts_with_equals": bool(starts_with_equals),
                "length_valid": bool(length_valid),
                "english_formula_style_expected": bool(english_style),
                "reference_candidates": references,
                "compatibility_warnings": warnings,
                "guarantee_level": "basic",
                "notes": notes,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def fill_range(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        formula: str,
        formula_type: str = "standard",
        preview_rows: int = 5,
    ) -> dict[str, Any]:
        if not formula.startswith("="):
            raise ExcelForgeError(ErrorCode.E400_FORMULA_INVALID, "Formula must start with '='")
        if len(formula) > 8192:
            raise ExcelForgeError(ErrorCode.E400_FORMULA_INVALID, "Formula exceeds max length")

        target_ref = parse_range(range_address)
        cell_count = target_ref.cell_count
        if cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(
                ErrorCode.E413_RANGE_TOO_LARGE,
                f"Formula fill range too large: {cell_count}",
            )

        def op(ctx: Any) -> dict[str, Any]:
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

            if bool(workbook.ReadOnly):
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")
            if bool(ws.ProtectContents):
                raise ExcelForgeError(ErrorCode.E403_SHEET_PROTECTED, "Worksheet is protected")

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=range_address,
                source_tool="formula.fill_range",
            )

            filled_formula = formula
            if "{row}" in formula:
                filled_formula = formula.replace("{row}", str(target_ref.start.row))

            try:
                rng = ws.Range(range_address)
                spill_range = None
                calculation_completed = True
                has_spill_error = False

                if formula_type == "dynamic_array":
                    app = ctx.app_manager.ensure_app()
                    supports_da = check_dynamic_array_support(app)
                    if not supports_da:
                        raise ExcelForgeError(
                            ErrorCode.E409_DYNAMIC_ARRAY_NOT_SUPPORTED,
                            "Excel version does not support dynamic arrays. Requires Excel 365 or 2021+.",
                        )
                    try:
                        rng.Formula2 = filled_formula
                    except AttributeError:
                        raise ExcelForgeError(
                            ErrorCode.E409_DYNAMIC_ARRAY_NOT_SUPPORTED,
                            "Dynamic arrays not supported in this Excel version.",
                        )

                    calculation_completed = wait_for_calculation(app, self._config.limits.calculation_timeout_seconds)

                    try:
                        if rng.HasSpill:
                            spill_range = rng.SpillingToRange.Address.replace("$", "")
                    except Exception:
                        pass

                    cell_value = rng.Value
                    if isinstance(cell_value, int) and cell_value in ERROR_VALUES:
                        has_spill_error = True
                elif formula_type == "array":
                    rng.FormulaArray = filled_formula
                    rng.Calculate()
                else:
                    if cell_count > 1 and "{row}" in formula:
                        row_start = target_ref.start.row
                        col_start = target_ref.start.col
                        for row_num in range(row_start, target_ref.end.row + 1):
                            cell_formula = formula.replace("{row}", str(row_num))
                            for col_num in range(col_start, target_ref.end.col + 1):
                                ws.Cells(row_num, col_num).Formula2 = cell_formula
                    elif cell_count > 1:
                        anchor = ws.Cells(target_ref.start.row, target_ref.start.col)
                        anchor.Formula2 = filled_formula
                        source_range = anchor
                        target_range = ws.Range(range_address)
                        try:
                            source_range.AutoFill(target_range)
                        except Exception:
                            row_start = target_ref.start.row
                            col_start = target_ref.start.col
                            for row_num in range(row_start, target_ref.end.row + 1):
                                for col_num in range(col_start, target_ref.end.col + 1):
                                    ws.Cells(row_num, col_num).Formula2 = filled_formula
                    else:
                        anchor = ws.Cells(target_ref.start.row, target_ref.start.col)
                        anchor.Formula2 = filled_formula
                    rng.Calculate()

                preview: list[dict[str, Any]] = []
                preview_end_row = min(target_ref.end.row, target_ref.start.row + preview_rows - 1)
                for row in range(target_ref.start.row, preview_end_row + 1):
                    for col in range(target_ref.start.col, target_ref.end.col + 1):
                        cell = ws.Cells(row, col)
                        preview.append(
                            {
                                "cell": str(cell.Address),
                                "formula": str(cell.Formula) if hasattr(cell, "Formula") else "",
                                "value": to_scalar(cell.Value2),
                            }
                        )
            except ExcelForgeError:
                raise
            except Exception as exc:
                import traceback
                error_detail = f"{exc}\n{traceback.format_exc()}"
                self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(
                    ErrorCode.E500_INTERNAL,
                    f"Failed to fill formula range: {error_detail}",
                ) from exc

            return {
                "sheet_name": sheet_name,
                "affected_range": range_address,
                "cells_written": cell_count,
                "formula_type": formula_type,
                "anchor_formula": formula,
                "preview": preview,
                "snapshot_id": snapshot_id,
                "spill_range": spill_range,
                "calculation_completed": calculation_completed,
                "has_spill_error": has_spill_error,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def set_single(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        cell: str,
        formula: str,
        formula_type: str = "standard",
    ) -> dict[str, Any]:
        if not formula.startswith("="):
            raise ExcelForgeError(ErrorCode.E400_FORMULA_INVALID, "Formula must start with '='")
        if len(formula) > 8192:
            raise ExcelForgeError(ErrorCode.E400_FORMULA_INVALID, "Formula exceeds max length")

        parsed = parse_cell_address(cell)

        def op(ctx: Any) -> dict[str, Any]:
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

            if bool(workbook.ReadOnly):
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")
            if bool(ws.ProtectContents):
                raise ExcelForgeError(ErrorCode.E403_SHEET_PROTECTED, "Worksheet is protected")

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=cell,
                source_tool="formula.set_single",
            )

            try:
                rng = ws.Cells(parsed.row, parsed.col)
                spill_range = None
                spill_preview: list[dict[str, Any]] = []
                calculation_completed = True
                has_error = False
                error_type: str | None = None

                if formula_type == "dynamic_array":
                    app = ctx.app_manager.ensure_app()
                    supports_da = check_dynamic_array_support(app)
                    if not supports_da:
                        raise ExcelForgeError(
                            ErrorCode.E409_DYNAMIC_ARRAY_NOT_SUPPORTED,
                            "Excel version does not support dynamic arrays. Requires Excel 365 or 2021+.",
                        )
                    try:
                        rng.Formula2 = formula
                    except AttributeError:
                        raise ExcelForgeError(
                            ErrorCode.E409_DYNAMIC_ARRAY_NOT_SUPPORTED,
                            "Dynamic arrays not supported in this Excel version.",
                        )

                    calculation_completed = wait_for_calculation(app, self._config.limits.calculation_timeout_seconds)

                    if rng.HasSpill:
                        try:
                            spill_rng = rng.SpillingToRange
                            spill_range = spill_rng.Address.replace("$", "")
                            max_preview = min(10, spill_rng.Cells.Count)
                            for i in range(1, max_preview + 1):
                                c = spill_rng.Cells(i)
                                spill_preview.append({
                                    "cell": str(c.Address),
                                    "value": to_scalar(c.Value2),
                                })
                        except Exception:
                            pass

                    cell_value = rng.Value
                    if isinstance(cell_value, int) and cell_value in ERROR_VALUES:
                        has_error = True
                        error_type = ERROR_VALUES.get(cell_value)
                else:
                    rng.Formula2 = formula
                    rng.Calculate()

                calculated_value = to_scalar(rng.Value2)
                cell_value = rng.Value
                if isinstance(cell_value, int) and cell_value in ERROR_VALUES:
                    has_error = True
                    error_type = ERROR_VALUES.get(cell_value)

            except ExcelForgeError:
                raise
            except Exception as exc:
                import traceback
                error_detail = f"{exc}\n{traceback.format_exc()}"
                self._restore_snapshot_best_effort(ctx, snapshot_id, workbook_id)
                raise ExcelForgeError(
                    ErrorCode.E500_INTERNAL,
                    f"Failed to set formula: {error_detail}",
                ) from exc

            return {
                "sheet_name": sheet_name,
                "cell": cell,
                "formula": formula,
                "formula_type": formula_type,
                "calculated_value": calculated_value,
                "has_error": has_error,
                "error_type": error_type,
                "spill_range": spill_range,
                "spill_preview": spill_preview,
                "calculation_completed": calculation_completed,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def get_dependencies(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        cell: str,
    ) -> dict[str, Any]:
        parsed = parse_cell_address(cell)

        CROSS_SHEET_REF_RE = re.compile(r"'?([^'!()=]+)'?!([A-Za-z]+:?[A-Za-z\d]*)")

        def op(ctx: Any) -> dict[str, Any]:
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

            rng = ws.Cells(parsed.row, parsed.col)
            has_formula = bool(rng.HasFormula)
            formula: str | None = None
            if has_formula:
                try:
                    formula = str(rng.Formula)
                except Exception:
                    formula = None

            calculated_value = to_scalar(rng.Value2)

            precedents: list[dict[str, str]] = []
            seen_precedents: set[tuple[str, str]] = set()

            try:
                for area in rng.Precedents.Areas:
                    ws_name = str(area.Worksheet.Name)
                    addr = area.Address.replace("$", "")
                    key = (ws_name, addr)
                    if key not in seen_precedents:
                        seen_precedents.add(key)
                        precedents.append({"sheet": ws_name, "range": addr})
            except Exception:
                pass

            if formula:
                matches = CROSS_SHEET_REF_RE.findall(formula)
                for match in matches:
                    ref_sheet = match[0]
                    ref_range = match[1]
                    key = (ref_sheet, ref_range)
                    if key not in seen_precedents:
                        seen_precedents.add(key)
                        precedents.append({"sheet": ref_sheet, "range": ref_range})

            dependents: list[dict[str, str]] = []
            seen_dependents: set[tuple[str, str]] = set()
            try:
                for area in rng.Dependents.Areas:
                    ws_name = str(area.Worksheet.Name)
                    addr = area.Address.replace("$", "")
                    key = (ws_name, addr)
                    if key not in seen_dependents:
                        seen_dependents.add(key)
                        dependents.append({"sheet": ws_name, "range": addr})
            except Exception:
                pass

            return {
                "sheet_name": sheet_name,
                "cell": cell,
                "has_formula": has_formula,
                "formula": formula,
                "calculated_value": calculated_value,
                "precedents": precedents,
                "dependents": dependents,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def scan_formulas(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
    ) -> dict[str, Any]:
        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_read_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Scan range too large: {target_ref.cell_count}")

        ERROR_MAP = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
            -2146826259: "#NAME?",
            -2146826252: "#NULL!",
            -2146826256: "#NUM!",
            -2146826265: "#REF!",
            -2146826273: "#VALUE!",
        }

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
            workbook = handle.workbook_obj
            try:
                ws = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            findings = []
            all_referenced_sheets = set()
            total_cells = target_ref.cell_count
            formula_cells = 0
            error_cells = 0
            ref_errors = 0
            name_errors = 0
            other_errors = 0

            rng = ws.Range(range_address)
            for row in range(target_ref.start.row, target_ref.end.row + 1):
                for col in range(target_ref.start.col, target_ref.end.col + 1):
                    cell = ws.Cells(row, col)
                    if not bool(cell.HasFormula):
                        continue

                    formula_cells += 1
                    formula = str(cell.Formula) if cell.Formula else ""
                    value = cell.Value

                    error_type = None
                    if isinstance(value, int) and value in ERROR_MAP:
                        error_type = ERROR_MAP[value]
                        error_cells += 1
                        if error_type == "#REF!":
                            ref_errors += 1
                        elif error_type == "#NAME?":
                            name_errors += 1
                        else:
                            other_errors += 1

                    sheet_refs = re.findall(r"'?([^'!]+)'?!", formula)
                    has_external = any("[" in ref for ref in sheet_refs)
                    for ref in sheet_refs:
                        if "[" not in ref:
                            all_referenced_sheets.add(ref)

                    findings.append({
                        "cell": cell.Address.replace("$", ""),
                        "formula": formula,
                        "error_type": error_type,
                        "has_external_ref": has_external,
                        "referenced_sheets": list(set(sheet_refs)),
                    })

            return {
                "action": "scan",
                "sheet_name": sheet_name,
                "scanned_range": range_address,
                "total_cells": total_cells,
                "formula_cells": formula_cells,
                "error_cells": error_cells,
                "findings": findings,
                "summary": {
                    "ref_errors": ref_errors,
                    "name_errors": name_errors,
                    "other_errors": other_errors,
                    "external_refs": len([f for f in findings if f["has_external_ref"]]),
                    "unique_referenced_sheets": list(all_referenced_sheets),
                },
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def repair_formulas(
        self,
        *,
        workbook_id: str,
        sheet_name: str,
        range_address: str,
        replacements: list[dict],
    ) -> dict[str, Any]:
        if not replacements:
            raise ExcelForgeError(ErrorCode.E400_REPLACEMENT_INVALID, "At least one replacement rule required")

        for rule in replacements:
            if "old_ref" not in rule or "new_ref" not in rule:
                raise ExcelForgeError(ErrorCode.E400_REPLACEMENT_INVALID, "Each replacement must have old_ref and new_ref")
            if not rule.get("old_ref"):
                raise ExcelForgeError(ErrorCode.E400_REPLACEMENT_INVALID, "old_ref cannot be empty")

        target_ref = parse_range(range_address)
        if target_ref.cell_count > self._config.limits.max_write_cells:
            raise ExcelForgeError(ErrorCode.E413_RANGE_TOO_LARGE, f"Repair range too large: {target_ref.cell_count}")

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")
            workbook = handle.workbook_obj
            try:
                ws = workbook.Worksheets(sheet_name)
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

            if bool(workbook.ReadOnly):
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=range_address,
                source_tool="formula.repair_references",
            )

            modifications = []
            cells_modified = 0
            cells_unchanged = 0

            try:
                rng = ws.Range(range_address)
                for row in range(target_ref.start.row, target_ref.end.row + 1):
                    for col in range(target_ref.start.col, target_ref.end.col + 1):
                        cell = ws.Cells(row, col)
                        if not bool(cell.HasFormula):
                            continue

                        old_formula = str(cell.Formula) if cell.Formula else ""
                        new_formula = old_formula

                        for rule in replacements:
                            new_formula = new_formula.replace(rule["old_ref"], rule["new_ref"])

                        if new_formula != old_formula:
                            cell.Formula2 = new_formula
                            cells_modified += 1
                            modifications.append({
                                "cell": cell.Address.replace("$", ""),
                                "old_formula": old_formula,
                                "new_formula": new_formula,
                                "new_value": to_scalar(cell.Value),
                                "still_has_error": isinstance(cell.Value, int) and cell.Value < 0,
                            })
                        else:
                            cells_unchanged += 1

            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to repair formulas: {exc}") from exc

            return {
                "action": "repair",
                "sheet_name": sheet_name,
                "repaired_range": range_address,
                "replacements_applied": len(replacements),
                "cells_modified": cells_modified,
                "cells_unchanged": cells_unchanged,
                "modifications": modifications,
                "snapshot_id": snapshot_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def _restore_snapshot_best_effort(self, ctx: Any, snapshot_id: str, workbook_id: str) -> None:
        try:
            meta, payload = self._snapshot_service.load_snapshot(snapshot_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                return
            ws = handle.workbook_obj.Worksheets(str(meta["sheet_name"]))
            self._snapshot_service.restore_snapshot(workbook=handle, worksheet=ws, snapshot_payload=payload)
        except Exception:
            return
