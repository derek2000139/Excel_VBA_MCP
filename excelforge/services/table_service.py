from __future__ import annotations

import logging
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.models.table_models import (
    TableColumnInfo,
    TableCreateData,
    TableCreateRequest,
    TableDeleteData,
    TableDeleteRequest,
    TableInfo,
    TableInspectData,
    TableInspectRequest,
    TableListData,
    TableListRequest,
    TableRenameData,
    TableRenameRequest,
    TableResizeData,
    TableResizeRequest,
    TableSetStyleData,
    TableSetStyleRequest,
    TableStyleInfo,
    TableToggleTotalRowData,
    TableToggleTotalRowRequest,
)
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.services.backup_service import BackupService
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.address_parser import parse_range

logger = logging.getLogger(__name__)


class TableService:
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

    def list_tables(self, params: TableListRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            tables: list[dict[str, Any]] = []

            if params.sheet_name:
                sheet_names = [params.sheet_name]
            else:
                sheet_names = [ws.Name for ws in workbook.Worksheets]

            for sname in sheet_names:
                try:
                    ws = workbook.Worksheets(sname)
                    for lst in ws.ListObjects:
                        table_info = self._extract_table_info(lst, sname)
                        tables.append(table_info.model_dump(mode="json") if hasattr(table_info, 'model_dump') else table_info)
                except Exception as exc:
                    logger.warning("Failed to list tables on sheet %s: %s", sname, exc)
                    continue

            return {"tables": tables, "total_count": len(tables)}

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def create_table(self, params: TableCreateRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            ws = self._get_worksheet(handle.workbook_obj, params.sheet_name)

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=ws,
                range_address=params.range_address,
                source_tool="table.create",
            )

            try:
                target_range = ws.Range(params.range_address)
                table_name = params.table_name or f"Table_{hash(params.range_address) & 0xFFFFFFFF}"

                lst = ws.ListObjects.Add(Source=target_range)
                lst.Name = table_name
                if params.style_name:
                    try:
                        lst.TableStyle = params.style_name
                    except Exception:
                        pass

                result = {
                    "name": lst.Name,
                    "sheet_name": ws.Name,
                    "range_address": lst.Range.Address,
                    "columns": self._get_table_columns(lst),
                    "total_row_count": int(lst.Range.Rows.Count),
                    "data_row_count": int(lst.DataBodyRange.Rows.Count) if lst.DataBodyRange else 0,
                    "has_header": params.has_header,
                    "has_total_row": bool(lst.ShowTotals),
                    "snapshot_id": snapshot_id,
                }
                try:
                    if lst.TableStyle:
                        result["style"] = {"name": str(lst.TableStyle.Name)}
                except Exception:
                    pass
                return result
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E430_TABLE_CREATE_FAILED, f"Failed to create table: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def inspect_table(self, params: TableInspectRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)

            result = {
                "name": lst.Name,
                "sheet_name": lst.Parent.Name,
                "range_address": lst.Range.Address,
                "columns": self._get_table_columns(lst),
                "total_row_count": int(lst.Range.Rows.Count),
                "data_row_count": int(lst.DataBodyRange.Rows.Count) if lst.DataBodyRange else 0,
                "has_header": bool(lst.HeaderRowRange),
                "has_total_row": bool(lst.ShowTotals),
            }
            try:
                if lst.TableStyle:
                    result["style"] = {"name": str(lst.TableStyle.Name)}
            except Exception:
                pass
            return result

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def resize_table(self, params: TableResizeRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)
            old_range_address = lst.Range.Address

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=lst.Parent,
                range_address=params.new_range_address,
                source_tool="table.resize",
            )

            try:
                new_range = lst.Parent.Range(params.new_range_address)
                lst.Resize(new_range)

                result = {
                    "name": lst.Name,
                    "old_range_address": old_range_address,
                    "new_range_address": lst.Range.Address,
                    "columns": self._get_table_columns(lst),
                    "total_row_count": int(lst.Range.Rows.Count),
                    "data_row_count": int(lst.DataBodyRange.Rows.Count) if lst.DataBodyRange else 0,
                    "has_total_row": bool(lst.ShowTotals),
                    "snapshot_id": snapshot_id,
                }
                return result
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to resize table: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def rename_table(self, params: TableRenameRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)
            old_name = lst.Name

            try:
                lst.Name = params.new_name
                return {
                    "old_name": old_name,
                    "new_name": lst.Name,
                    "sheet_name": lst.Parent.Name,
                    "range_address": lst.Range.Address,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to rename table: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def set_table_style(self, params: TableSetStyleRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)
            old_style = None
            try:
                if lst.TableStyle:
                    old_style = {"name": str(lst.TableStyle.Name)}
            except Exception:
                pass

            try:
                if params.style_name:
                    lst.TableStyle = params.style_name
                else:
                    lst.TableStyle = None
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to set table style: {exc}") from exc

            new_style = None
            try:
                if lst.TableStyle:
                    new_style = {"name": str(lst.TableStyle.Name)}
            except Exception:
                pass
                return {
                    "name": lst.Name,
                    "old_style": old_style,
                    "new_style": new_style,
                    "sheet_name": lst.Parent.Name,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to set table style: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def toggle_total_row(self, params: TableToggleTotalRowRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=lst.Parent,
                range_address=lst.Range.Address,
                source_tool="table.toggle_total_row",
            )

            try:
                lst.ShowTotals = not lst.ShowTotals
                return {
                    "name": lst.Name,
                    "sheet_name": lst.Parent.Name,
                    "has_total_row": bool(lst.ShowTotals),
                    "snapshot_id": snapshot_id,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to toggle total row: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def delete_table(self, params: TableDeleteRequest) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(params.workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {params.workbook_id}")

            if handle.workbook_obj.ReadOnly:
                raise ExcelForgeError(ErrorCode.E409_WORKBOOK_READONLY, "Workbook is read-only")

            workbook = handle.workbook_obj
            lst = self._find_table(workbook, params.table_name, params.sheet_name)
            data_range_address = lst.Range.Address
            sheet_name = lst.Parent.Name
            table_name = lst.Name

            snapshot_id = self._snapshot_service.create_snapshot(
                workbook=handle,
                worksheet=lst.Parent,
                range_address=data_range_address,
                source_tool="table.delete",
            )

            try:
                lst.Delete()
                return {
                    "name": table_name,
                    "sheet_name": sheet_name,
                    "data_preserved": True,
                    "data_range_address": data_range_address,
                    "snapshot_id": snapshot_id,
                }
            except Exception as exc:
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to delete table: {exc}") from exc

        return self._worker.submit(op, timeout_seconds=self._config.limits.operation_timeout_seconds, requires_excel=True)

    def _find_table(self, workbook: Any, table_name: str, sheet_name: str | None) -> Any:
        if sheet_name:
            ws = workbook.Worksheets(sheet_name)
        else:
            for ws in workbook.Worksheets:
                for lst in ws.ListObjects:
                    if lst.Name == table_name:
                        return lst
            raise ExcelForgeError(ErrorCode.E430_TABLE_NOT_FOUND, f"Table not found: {table_name}")

        try:
            return ws.ListObjects(table_name)
        except Exception:
            raise ExcelForgeError(ErrorCode.E430_TABLE_NOT_FOUND, f"Table not found: {table_name} on sheet {sheet_name}")

    def _get_worksheet(self, workbook: Any, sheet_name: str) -> Any:
        try:
            return workbook.Worksheets(sheet_name)
        except Exception as exc:
            raise ExcelForgeError(ErrorCode.E404_SHEET_NOT_FOUND, f"Sheet not found: {sheet_name}") from exc

    def _extract_table_info(self, lst: Any, sheet_name: str) -> dict[str, Any]:
        columns = self._get_table_columns(lst)
        result = {
            "name": lst.Name,
            "sheet_name": sheet_name,
            "range_address": lst.Range.Address,
            "columns": columns,
            "total_row_count": int(lst.Range.Rows.Count),
            "data_row_count": int(lst.DataBodyRange.Rows.Count) if lst.DataBodyRange else 0,
            "has_total_row": bool(lst.ShowTotals),
        }
        try:
            if lst.TableStyle:
                result["style"] = {"name": str(lst.TableStyle.Name)}
        except Exception:
            pass
        return result

    def _get_table_columns(self, lst: Any) -> list[dict[str, Any]]:
        columns: list[dict[str, Any]] = []
        try:
            if lst.HeaderRowRange:
                for idx, cell in enumerate(lst.HeaderRowRange.Cells, start=1):
                    columns.append({
                        "name": str(cell.Value) if cell.Value else f"Column{idx}",
                        "index": idx,
                        "field_type": "unknown",
                    })
        except Exception:
            pass
        return columns
