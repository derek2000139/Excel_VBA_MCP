from __future__ import annotations

from typing import Any, Callable

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime_api.audit_api import AuditApi
from excelforge.runtime_api.context import RuntimeApiContext
from excelforge.runtime_api.formula_api import FormulaApi
from excelforge.runtime_api.format_api import FormatApi
from excelforge.runtime_api.names_api import NamesApi
from excelforge.runtime_api.pq_api import PqApi
from excelforge.runtime_api.range_api import RangeApi
from excelforge.runtime_api.recovery_api import RecoveryApi
from excelforge.runtime_api.server_api import ServerApi
from excelforge.runtime_api.sheet_api import SheetApi
from excelforge.runtime_api.vba_api import VbaApi
from excelforge.runtime_api.workbook_api import WorkbookApi

MethodFn = Callable[[dict[str, Any], str], dict[str, Any]]

_NO_EXCEL_REQUIRED = frozenset({
    "server.health",
})


class RuntimeApiDispatcher:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx
        self._config = ctx.services.config
        workbook = WorkbookApi(ctx)
        sheet = SheetApi(ctx)
        rng = RangeApi(ctx)
        formula = FormulaApi(ctx)
        fmt = FormatApi(ctx)
        vba = VbaApi(ctx)
        names = NamesApi(ctx)
        recovery = RecoveryApi(ctx)
        audit = AuditApi(ctx)
        server = ServerApi(ctx)
        pq = PqApi(ctx)

        self._methods: dict[str, MethodFn] = {
            "workbook.open": workbook.open,
            "workbook.close": workbook.close,
            "workbook.create": workbook.create,
            "workbook.save": workbook.save,
            "workbook.list": workbook.list,
            "workbook.info": workbook.info,
            "sheet.inspect": sheet.inspect,
            "sheet.create": sheet.create,
            "sheet.rename": sheet.rename,
            "sheet.preview_delete": sheet.preview_delete,
            "sheet.delete": sheet.delete,
            "sheet.auto_filter": sheet.auto_filter,
            "sheet.get_conditional_formats": sheet.get_conditional_formats,
            "sheet.get_data_validations": sheet.get_data_validations,
            "range.read": rng.read,
            "range.write": rng.write,
            "range.clear": rng.clear,
            "range.copy": rng.copy,
            "range.insert_rows": rng.insert_rows,
            "range.delete_rows": rng.delete_rows,
            "range.insert_columns": rng.insert_columns,
            "range.delete_columns": rng.delete_columns,
            "range.sort": rng.sort,
            "range.merge": rng.merge,
            "range.unmerge": rng.unmerge,
            "formula.fill": formula.fill,
            "formula.set_single": formula.set_single,
            "formula.get_dependencies": formula.get_dependencies,
            "formula.repair": formula.repair,
            "format.set_style": fmt.set_style,
            "format.auto_fit": fmt.auto_fit,
            "vba.inspect_project": vba.inspect_project,
            "vba.get_module_code": vba.get_module_code,
            "vba.scan_code": vba.scan_code,
            "vba.sync_module": vba.sync_module,
            "vba.remove_module": vba.remove_module,
            "vba.execute_macro": vba.execute_macro,
            "vba.import_module": vba.import_module,
            "vba.export_module": vba.export_module,
            "vba.compile": vba.compile,
            "names.list": names.list,
            "names.read": names.read,
            "names.create": names.create,
            "names.delete": names.delete,
            "recovery.list_snapshots": recovery.list_snapshots,
            "recovery.preview_snapshot": recovery.preview_snapshot,
            "recovery.restore_snapshot": recovery.restore_snapshot,
            "recovery.undo_last": recovery.undo_last,
            "recovery.snapshot_stats": recovery.snapshot_stats,
            "recovery.snapshot_cleanup": recovery.snapshot_cleanup,
            "recovery.list_backups": recovery.list_backups,
            "recovery.restore_backup": recovery.restore_backup,
            "audit.list_operations": audit.list_operations,
            "server.status": server.status,
            "server.health": server.health,
            "pq.list_queries": pq.list_queries,
            "pq.get_query_code": pq.get_query_code,
            "pq.update_query": pq.update_query,
            "pq.refresh": pq.refresh,
            "pq.list_connections": pq.list_connections,
        }

    def dispatch(self, method: str, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        handler = self._methods.get(method)
        if handler is None:
            raise ExcelForgeError(
                ErrorCode.E400_BAD_REQUEST,
                f"Unsupported runtime method: {method}",
            )

        if method not in _NO_EXCEL_REQUIRED:
            self._check_ready()

        return handler(params, actor_id)

    def _check_ready(self) -> None:
        worker = self._ctx.services.worker
        if worker.is_ready():
            return

        wait_timeout = self._config.excel.request_wait_ready_seconds
        ready = worker.wait_ready(timeout=wait_timeout)
        if ready:
            return

        ready_status = worker.get_ready_status()
        if ready_status["warmup_error"]:
            raise ExcelForgeError(
                ErrorCode.E503_EXCEL_INIT_FAILED,
                f"Excel initialization failed: {ready_status['warmup_error']}",
            )
        else:
            raise ExcelForgeError(
                ErrorCode.E503_RUNTIME_NOT_READY,
                "Excel engine is still initializing, please retry shortly",
            )

    def method_names(self) -> list[str]:
        return sorted(self._methods.keys())
