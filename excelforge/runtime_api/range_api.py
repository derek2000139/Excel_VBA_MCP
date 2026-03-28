from __future__ import annotations

from typing import Any

from excelforge.models.range_models import (
    RangeClearContentsRequest,
    RangeCopyRangeRequest,
    RangeDeleteColumnsRequest,
    RangeDeleteRowsRequest,
    RangeInsertColumnsRequest,
    RangeInsertRowsRequest,
    RangeMergeCellsRequest,
    RangeReadValuesRequest,
    RangeSortDataRequest,
    RangeUnmergeCellsRequest,
    RangeWriteValuesRequest,
)
from excelforge.runtime_api.context import RuntimeApiContext


class RangeApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def read(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeReadValuesRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            value_mode=params.get("value_mode", "raw"),
            include_formulas=bool(params.get("include_formulas", False)),
            row_offset=int(params.get("row_offset", 0)),
            row_limit=int(params.get("row_limit", 200)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.read",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.read_values(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                value_mode=req.value_mode,
                include_formulas=req.include_formulas,
                row_offset=req.row_offset,
                row_limit=req.row_limit,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "value_mode": req.value_mode,
                "include_formulas": req.include_formulas,
                "row_offset": req.row_offset,
                "row_limit": req.row_limit,
            },
            default_workbook_id=req.workbook_id,
        )

    def write(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeWriteValuesRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            start_cell=params.get("start_cell") or params.get("range", ""),
            values=params.get("values") or [],
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.write",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.write_values(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                start_cell=req.start_cell,
                values=req.values,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "start_cell": req.start_cell},
            default_workbook_id=req.workbook_id,
        )

    def clear(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeClearContentsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            scope=params.get("scope", "contents"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.clear",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.clear_contents(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                scope=req.scope,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "scope": req.scope,
            },
            default_workbook_id=req.workbook_id,
        )

    def copy(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeCopyRangeRequest(
            workbook_id=params.get("workbook_id", ""),
            source_sheet=params.get("sheet_name") or params.get("source_sheet", ""),
            source_range=params.get("source_range", ""),
            target_sheet=params.get("target_sheet", ""),
            target_start_cell=params.get("target_start_cell") or params.get("target_range", ""),
            paste_mode=params.get("paste_mode", "values"),
            target_workbook_id=params.get("target_workbook_id"),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.copy",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.copy_range(
                workbook_id=req.workbook_id,
                source_sheet=req.source_sheet,
                source_range=req.source_range,
                target_sheet=req.target_sheet,
                target_start_cell=req.target_start_cell,
                paste_mode=req.paste_mode,
                target_workbook_id=req.target_workbook_id,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "source_sheet": req.source_sheet,
                "source_range": req.source_range,
                "target_sheet": req.target_sheet,
                "target_start_cell": req.target_start_cell,
                "paste_mode": req.paste_mode,
                "target_workbook_id": req.target_workbook_id,
            },
            default_workbook_id=req.workbook_id,
        )

    def insert_rows(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeInsertRowsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            row_number=int(params.get("row", params.get("row_number", 1))),
            count=int(params.get("count", 1)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.insert_rows",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.insert_rows(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                row_number=req.row_number,
                count=req.count,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "row_number": req.row_number, "count": req.count},
            default_workbook_id=req.workbook_id,
        )

    def delete_rows(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeDeleteRowsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            start_row=int(params.get("row", params.get("start_row", 1))),
            count=int(params.get("count", 1)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.delete_rows",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.delete_rows(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                start_row=req.start_row,
                count=req.count,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "start_row": req.start_row, "count": req.count},
            default_workbook_id=req.workbook_id,
        )

    def insert_columns(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeInsertColumnsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            column=params.get("column", ""),
            count=int(params.get("count", 1)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.insert_columns",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.insert_columns(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                column=req.column,
                count=req.count,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "column": req.column, "count": req.count},
            default_workbook_id=req.workbook_id,
        )

    def delete_columns(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeDeleteColumnsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            start_column=params.get("column", params.get("start_column", "")),
            count=int(params.get("count", 1)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.delete_columns",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.delete_columns(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                start_column=req.start_column,
                count=req.count,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "start_column": req.start_column, "count": req.count},
            default_workbook_id=req.workbook_id,
        )

    def sort(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeSortDataRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            sort_fields=params.get("sort_keys") or params.get("sort_fields") or [],
            has_header=bool(params.get("has_header", False)),
            case_sensitive=bool(params.get("case_sensitive", False)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.sort",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.sort_data(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                sort_fields=req.sort_fields,
                has_header=req.has_header,
                case_sensitive=req.case_sensitive,
            ),
            args_summary={
                "workbook_id": req.workbook_id,
                "sheet_name": req.sheet_name,
                "range": req.range,
                "sort_fields": req.sort_fields,
                "has_header": req.has_header,
                "case_sensitive": req.case_sensitive,
            },
            default_workbook_id=req.workbook_id,
        )

    def merge(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeMergeCellsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            across=bool(params.get("across", False)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.merge",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.merge_cells(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
                across=req.across,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "range": req.range, "across": req.across},
            default_workbook_id=req.workbook_id,
        )

    def unmerge(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = RangeUnmergeCellsRequest(
            workbook_id=params.get("workbook_id", ""),
            sheet_name=params.get("sheet_name", ""),
            range=params.get("range", ""),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="range.unmerge",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.range_service.unmerge_cells(
                workbook_id=req.workbook_id,
                sheet_name=req.sheet_name,
                range_address=req.range,
            ),
            args_summary={"workbook_id": req.workbook_id, "sheet_name": req.sheet_name, "range": req.range},
            default_workbook_id=req.workbook_id,
        )
