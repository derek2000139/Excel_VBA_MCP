from __future__ import annotations

from typing import Any

from excelforge.models.formula_models import (
    FormulaFillRangeRequest,
    FormulaGetDependenciesRequest,
    FormulaRepairReferencesRequest,
    FormulaSetSingleRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class FormulaApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def fill(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = FormulaFillRangeRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            formula=params.get("formula", ""),
            formula_type=params.get("formula_type", "standard"),
            preview_rows=int(params.get("preview_rows", 5)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="formula.fill",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.formula_service.fill_range(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                formula=req.formula,
                formula_type=req.formula_type,
                preview_rows=req.preview_rows,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "formula_type": req.formula_type,
            },
            default_workbook_id=req.workbook_id,
        )

    def set_single(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = FormulaSetSingleRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            cell=params.get("cell", ""),
            formula=params.get("formula", ""),
            formula_type=params.get("formula_type", "standard"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="formula.set_single",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.formula_service.set_single(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                cell=req.cell,
                formula=req.formula,
                formula_type=req.formula_type,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "cell": req.cell},
            default_workbook_id=req.workbook_id,
        )

    def get_dependencies(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = FormulaGetDependenciesRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            cell=params.get("cell", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="formula.get_dependencies",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.formula_service.get_dependencies(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                cell=req.cell,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "cell": req.cell},
            default_workbook_id=req.workbook_id,
        )

    def repair(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = FormulaRepairReferencesRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            action=params.get("action", "scan"),
            replacements=params.get("replacements"),
            client_request_id=params.get("client_request_id"),
        )
        if req.action == "scan":
            op = lambda: self._ctx.services.formula_service.scan_formulas(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
            )
        else:
            op = lambda: self._ctx.services.formula_service.repair_formulas(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                replacements=req.replacements or [],
            )
        return self._ctx.run_operation(
            method_name="formula.repair",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=op,
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "action": req.action,
            },
            default_workbook_id=req.workbook_id,
        )
