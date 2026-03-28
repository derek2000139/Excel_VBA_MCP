from __future__ import annotations

from typing import Any

from excelforge.models.workbook_models import WorkbookCreateFileRequest, WorkbookOpenRequest, WorkbookSaveFileRequest
from excelforge.runtime_api.context import RuntimeApiContext


class WorkbookApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def open(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookOpenRequest(
            file_path=params.get("file_path", ""),
            read_only=bool(params.get("read_only", False)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.open",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_service.open_file(req.file_path, req.read_only),
            args_summary={"file_path": req.file_path, "read_only": req.read_only},
            default_file_path=req.file_path,
        )

    def close(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        save_before_close = bool(params.get("save_before_close", False))

        def op() -> dict[str, Any]:
            if save_before_close:
                self._ctx.services.workbook_service.save_file(workbook_id=workbook_id)
            return self._ctx.services.workbook_service.close_file(
                workbook_id=workbook_id,
                force_discard=not save_before_close,
            )

        return self._ctx.run_operation(
            method_name="workbook.close",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=op,
            args_summary={"workbook_id": workbook_id, "save_before_close": save_before_close},
            default_workbook_id=workbook_id,
        )

    def create(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookCreateFileRequest(
            file_path=params.get("file_path", ""),
            sheet_names=params.get("sheet_names") or ["Sheet1"],
            overwrite=bool(params.get("overwrite", False)),
            client_request_id=params.get("client_request_id"),
        )
        return self._ctx.run_operation(
            method_name="workbook.create",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_service.create_file(
                file_path=req.file_path,
                sheet_names=req.sheet_names,
                overwrite=req.overwrite,
            ),
            args_summary={
                "file_path": req.file_path,
                "sheet_names": req.sheet_names,
                "overwrite": req.overwrite,
            },
            default_file_path=req.file_path,
        )

    def save(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        req = WorkbookSaveFileRequest(
            workbook_id=params.get("workbook_id", ""),
            save_as_path=params.get("save_as_path"),
            client_request_id=params.get("client_request_id"),
        )
        save_as_path = req.save_as_path or ""
        return self._ctx.run_operation(
            method_name="workbook.save",
            actor_id=actor_id,
            client_request_id=req.client_request_id,
            operation_fn=lambda: self._ctx.services.workbook_service.save_file(
                workbook_id=req.workbook_id,
                save_as_path=save_as_path,
            ),
            args_summary={"workbook_id": req.workbook_id, "save_as_path": save_as_path},
            default_workbook_id=req.workbook_id,
        )

    def list(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        return self._ctx.run_operation(
            method_name="workbook.list",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.workbook_service.list_open(),
            args_summary={},
        )

    def info(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        return self._ctx.run_operation(
            method_name="workbook.info",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.workbook_service.get_info(workbook_id),
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )
