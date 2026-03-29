from __future__ import annotations

from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.runtime.workbook_registry import WorkbookHandle
from excelforge.services.snapshot_service import SnapshotService
from excelforge.utils.address_parser import range_to_a1
from excelforge.utils.file_format import (
    EXTENSION_FORMAT_MAP,
    get_file_format,
    supports_vba,
    validate_extension_for_save,
)
from excelforge.utils.ids import generate_workbook_id
from excelforge.utils.path_guard import ensure_same_extension, normalize_allowed_path
from excelforge.utils.timestamps import utc_now_rfc3339


class WorkbookService:
    def __init__(
        self,
        config: AppConfig,
        worker: ExcelWorker,
        snapshot_service: SnapshotService,
    ) -> None:
        self._config = config
        self._worker = worker
        self._snapshot_service = snapshot_service

    def open_file(self, file_path: str, read_only: bool) -> dict[str, Any]:
        normalized = normalize_allowed_path(file_path, [Path(p) for p in self._config.paths.allowed_roots], set(self._config.paths.allowed_extensions))

        def op(ctx: Any) -> dict[str, Any]:
            existing = ctx.registry.find_by_path(str(normalized))
            if existing is not None:
                workbook = existing.workbook_obj
                try:
                    _ = workbook.Name
                except Exception:
                    ctx.registry.remove(existing.workbook_id)
                    existing = None

            if existing is not None:
                workbook = existing.workbook_obj
                existing_read_only = bool(existing.read_only or workbook.ReadOnly)
                if existing_read_only and not read_only:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception as exc:  # noqa: BLE001
                        raise ExcelForgeError(
                            ErrorCode.E409_WORKBOOK_READONLY,
                            f"Workbook is read-only and cannot be reopened for write: {exc}",
                        ) from exc
                    ctx.registry.remove(existing.workbook_id)
                    existing = None
                else:
                    return {
                        "workbook_id": existing.workbook_id,
                        "workbook_name": existing.workbook_name,
                        "file_path": existing.file_path,
                        "read_only": bool(workbook.ReadOnly),
                        "already_open": True,
                        "has_macros": self._has_macros(existing.file_path, workbook),
                        "sheet_names": [str(s.Name) for s in workbook.Worksheets],
                        "opened_at": existing.opened_at,
                        "file_format": existing.file_format,
                        "max_rows": existing.max_rows,
                        "max_columns": existing.max_columns,
                        "vba_enabled": supports_vba(Path(existing.file_path).suffix.lower()),
                    }

            if ctx.registry.count() >= self._config.limits.max_open_workbooks:
                raise ExcelForgeError(
                    ErrorCode.E423_FEATURE_NOT_SUPPORTED,
                    f"max_open_workbooks exceeded: {self._config.limits.max_open_workbooks}",
                )

            try:
                workbook = ctx.app_manager.open_workbook(normalized, read_only)
            except ExcelForgeError:
                raise
            except Exception as exc:  # noqa: BLE001
                raise ExcelForgeError(ErrorCode.E409_FILE_LOCKED, f"Failed to open workbook: {exc}") from exc

            actual_read_only = bool(workbook.ReadOnly)
            if not read_only and actual_read_only:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_READONLY,
                    "Workbook can only be opened as read-only (likely locked by another process)",
                )

            workbook_id = generate_workbook_id(
                ctx.registry.generation,
                ctx.registry.runtime_fingerprint,
            )
            file_format, max_rows, max_columns = self._detect_file_format(str(normalized))
            handle = WorkbookHandle(
                workbook_id=workbook_id,
                workbook_name=str(workbook.Name),
                file_path=str(normalized),
                read_only=actual_read_only,
                opened_at=utc_now_rfc3339(),
                workbook_obj=workbook,
                file_format=file_format,
                max_rows=max_rows,
                max_columns=max_columns,
            )
            ctx.registry.add(handle)

            try:
                app = ctx.app_manager._app
                if app and not app.Visible:
                    app.Visible = True
                    app.ScreenUpdating = True
            except Exception:
                pass

            return {
                "workbook_id": workbook_id,
                "workbook_name": str(workbook.Name),
                "file_path": str(normalized),
                "read_only": actual_read_only,
                "already_open": False,
                "has_macros": self._has_macros(str(normalized), workbook),
                "sheet_names": [str(s.Name) for s in workbook.Worksheets],
                "opened_at": handle.opened_at,
                "file_format": file_format,
                "max_rows": max_rows,
                "max_columns": max_columns,
                "vba_enabled": supports_vba(Path(normalized).suffix.lower()),
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            allow_rebuild=True,
            requires_excel=True,
        )

    def list_open(self) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            items: list[dict[str, Any]] = []
            for handle in ctx.registry.list_items():
                workbook = handle.workbook_obj
                dirty = not bool(workbook.Saved)
                items.append(
                    {
                        "workbook_id": handle.workbook_id,
                        "workbook_name": handle.workbook_name,
                        "file_path": handle.file_path,
                        "read_only": bool(workbook.ReadOnly),
                        "dirty": dirty,
                        "stale": False,
                        "opened_at": handle.opened_at,
                    }
                )
            return {"items": items}

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def get_info(self, workbook_id: str) -> dict[str, Any]:
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
            sheets: list[dict[str, Any]] = []
            for idx, ws in enumerate(workbook.Worksheets, start=1):
                used = ws.UsedRange
                used_rows = int(used.Rows.Count)
                used_cols = int(used.Columns.Count)
                sheets.append(
                    {
                        "index": idx,
                        "name": str(ws.Name),
                        "visible": bool(ws.Visible != 0),
                        "protected": bool(ws.ProtectContents),
                        "used_range": range_to_a1_from_com(used),
                        "used_rows": used_rows,
                        "used_columns": used_cols,
                    }
                )
            return {
                "workbook_id": handle.workbook_id,
                "workbook_name": handle.workbook_name,
                "file_path": handle.file_path,
                "read_only": bool(workbook.ReadOnly),
                "dirty": not bool(workbook.Saved),
                "has_macros": self._has_macros(handle.file_path, workbook),
                "sheet_count": len(sheets),
                "active_sheet": str(workbook.ActiveSheet.Name),
                "sheets": sheets,
                "file_format": handle.file_format,
                "max_rows": handle.max_rows,
                "max_columns": handle.max_columns,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def save_file(self, workbook_id: str, save_as_path: str = "") -> dict[str, Any]:
        normalized_save_as: Path | None = None

        if save_as_path:
            normalized_save_as = normalize_allowed_path(save_as_path, self._config.paths.allowed_roots, set(self._config.paths.allowed_extensions))
            new_suffix = normalized_save_as.suffix.lower()
            new_format = get_file_format(new_suffix)
            if new_format is None:
                raise ExcelForgeError(
                    ErrorCode.E400_SAVE_AS_FORMAT_MISMATCH,
                    f"Cannot save as {new_suffix}. Supported: .xlsx, .xlsm, .xlsb, .xls",
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

            if normalized_save_as is None:
                workbook.Save()
                saved_path = Path(handle.file_path)
                save_type = "save"
                format_converted = False
                vba_stripped = False
                original_path = None
            else:
                original_ext = Path(handle.file_path).suffix.lower()
                new_ext = normalized_save_as.suffix.lower()
                vba_stripped, format_converted = validate_extension_for_save(original_ext, new_ext)

                workbook.SaveAs(str(normalized_save_as))
                original_path = handle.file_path
                handle.file_path = str(normalized_save_as)
                saved_path = normalized_save_as
                save_type = "save_as"

            return {
                "workbook_id": handle.workbook_id,
                "saved_path": str(saved_path),
                "dirty": not bool(workbook.Saved),
                "saved_at": utc_now_rfc3339(),
                "save_type": save_type,
                "format_converted": format_converted,
                "vba_stripped": vba_stripped,
                "original_path": original_path,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def close_file(self, workbook_id: str, force_discard: bool = False) -> dict[str, Any]:
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
            dirty = not bool(workbook.Saved)
            if dirty and not force_discard:
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_DIRTY,
                    "Workbook has unsaved changes and cannot be closed",
                )

            workbook.Close(SaveChanges=False)
            ctx.registry.remove(workbook_id)
            invalidated = self._snapshot_service.expire_workbook_snapshots(workbook_id)
            self._quit_if_no_workbooks(ctx)
            return {
                "workbook_id": workbook_id,
                "closed": True,
                "changes_discarded": bool(dirty and force_discard),
                "invalidated_snapshot_count": invalidated,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def _quit_if_no_workbooks(self, ctx: Any) -> None:
        """如果所有工作簿都已关闭，则退出 Excel。"""
        try:
            count_info = ctx.registry.get_workbook_count()
            registry_count = count_info.get("registry", 0)
            if registry_count == 0:
                app = ctx.app_manager._app
                if app is not None:
                    try:
                        app.Visible = False
                        app.Quit()
                        ctx.app_manager.invalidate()
                    except Exception:
                        pass
        except Exception:
            pass

    def create_file(self, file_path: str, sheet_names: list[str], overwrite: bool) -> dict[str, Any]:
        normalized = normalize_allowed_path(file_path, [Path(p) for p in self._config.paths.allowed_roots], set(self._config.paths.allowed_extensions))
        suffix = normalized.suffix.lower()
        file_format = get_file_format(suffix)
        if file_format is None:
            raise ExcelForgeError(
                ErrorCode.E400_UNSUPPORTED_EXTENSION,
                f"workbook.create_file does not support {suffix}. Supported: .xlsx, .xlsm, .xlsb, .xls",
            )

        if not sheet_names:
            sheet_names = ["Sheet1"]
        if len(sheet_names) > self._config.limits.max_create_sheets:
            raise ExcelForgeError(
                ErrorCode.E423_FEATURE_NOT_SUPPORTED,
                f"sheet_names exceeds max_create_sheets={self._config.limits.max_create_sheets}",
            )

        seen: set[str] = set()
        for name in sheet_names:
            self._validate_sheet_name(name)
            key = name.lower()
            if key in seen:
                raise ExcelForgeError(ErrorCode.E409_SHEET_NAME_EXISTS, f"Duplicate sheet name: {name}")
            seen.add(key)

        if normalized.exists() and not overwrite:
            raise ExcelForgeError(
                ErrorCode.E409_FILE_EXISTS,
                f"Target file already exists: {normalized}",
            )
        was_existing = normalized.exists()

        def op(ctx: Any) -> dict[str, Any]:
            if ctx.registry.count() >= self._config.limits.max_open_workbooks:
                raise ExcelForgeError(
                    ErrorCode.E423_FEATURE_NOT_SUPPORTED,
                    f"max_open_workbooks exceeded: {self._config.limits.max_open_workbooks}",
                )

            if normalized.exists() and overwrite:
                normalized.unlink(missing_ok=True)

            app = ctx.app_manager.ensure_app()
            wb = app.Workbooks.Add()
            try:
                while wb.Sheets.Count > 1:
                    wb.Sheets(wb.Sheets.Count).Delete()

                wb.Sheets(1).Name = sheet_names[0]
                for name in sheet_names[1:]:
                    ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                    ws.Name = name

                wb.SaveAs(str(normalized), FileFormat=file_format)
            except Exception as exc:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
                raise ExcelForgeError(ErrorCode.E500_INTERNAL, f"Failed to create workbook: {exc}") from exc

            workbook_id = generate_workbook_id(
                ctx.registry.generation,
                ctx.registry.runtime_fingerprint,
            )
            vba_enabled = supports_vba(suffix)
            handle = WorkbookHandle(
                workbook_id=workbook_id,
                workbook_name=str(wb.Name),
                file_path=str(normalized),
                read_only=False,
                opened_at=utc_now_rfc3339(),
                workbook_obj=wb,
                file_format=suffix.lstrip("."),
                max_rows=1048576,
                max_columns=16384,
            )
            ctx.registry.add(handle)
            return {
                "workbook_id": workbook_id,
                "workbook_name": str(wb.Name),
                "file_path": str(normalized),
                "sheet_names": [str(ws.Name) for ws in wb.Worksheets],
                "overwritten": bool(overwrite and was_existing),
                "created_at": utc_now_rfc3339(),
                "vba_enabled": vba_enabled,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            allow_rebuild=True,
            requires_excel=True,
        )

    @staticmethod
    def _has_macros(file_path: str, workbook: Any) -> bool:
        suffix = Path(file_path).suffix.lower()
        if suffix == ".xlsm":
            return True
        try:
            return bool(workbook.HasVBProject)
        except Exception:
            return False

    @staticmethod
    def _detect_file_format(file_path: str) -> tuple[str, int, int]:
        suffix = Path(file_path).suffix.lower()
        if suffix == ".xls":
            return ("xls", 65536, 256)
        elif suffix == ".xlsx":
            return ("xlsx", 1048576, 16384)
        elif suffix == ".xlsm":
            return ("xlsm", 1048576, 16384)
        elif suffix == ".xlsb":
            return ("xlsb", 1048576, 16384)
        else:
            return ("xlsx", 1048576, 16384)

    @staticmethod
    def _validate_sheet_name(name: str) -> None:
        invalid_chars = set("\\/?*[]")
        if not name or len(name) > 31:
            raise ExcelForgeError(ErrorCode.E400_INVALID_SHEET_NAME, f"Invalid sheet name length: {name}")
        if any(ch in invalid_chars for ch in name):
            raise ExcelForgeError(ErrorCode.E400_INVALID_SHEET_NAME, f"Invalid sheet name characters: {name}")

    def inspect(self, action: str, workbook_id: str = "") -> dict[str, Any]:
        if action == "list":
            return self.list_open()
        elif action == "info":
            if not workbook_id:
                raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, "workbook_id required for info action")
            return self.get_info(workbook_id)
        else:
            raise ExcelForgeError(ErrorCode.E400_INVALID_ARGUMENT, f"Invalid action: {action}")


def range_to_a1_from_com(range_obj: Any) -> str:
    start_row = int(range_obj.Row)
    start_col = int(range_obj.Column)
    rows = int(range_obj.Rows.Count)
    cols = int(range_obj.Columns.Count)

    from excelforge.utils.address_parser import CellRef, RangeRef

    rr = RangeRef(
        start=CellRef(start_row, start_col),
        end=CellRef(start_row + rows - 1, start_col + cols - 1),
    )
    return range_to_a1(rr)
