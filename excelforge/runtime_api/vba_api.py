from __future__ import annotations

from typing import Any

from excelforge.runtime_api.context import RuntimeApiContext


class VbaApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def inspect_project(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        return self._ctx.run_operation(
            method_name="vba.inspect_project",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.inspect_project(workbook_id=workbook_id),
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )

    def get_module_code(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        module_name = str(params.get("module_name", ""))
        return self._ctx.run_operation(
            method_name="vba.get_module_code",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.get_module_code(
                workbook_id=workbook_id,
                module_name=module_name,
            ),
            args_summary={"workbook_id": workbook_id, "module_name": module_name},
            default_workbook_id=workbook_id,
        )

    def scan_code(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        code = str(params.get("code", ""))
        module_name = params.get("module_name")
        module_type = str(params.get("module_type", "standard_module"))
        return self._ctx.run_operation(
            method_name="vba.scan_code",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.scan_code(
                code=code,
                module_name=module_name,
                module_type=module_type,
            ),
            args_summary={"module_name": module_name, "module_type": module_type},
        )

    def sync_module(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        module_name = str(params.get("module_name", ""))
        module_type = str(params.get("module_type", "standard_module"))
        code = str(params.get("code", ""))
        overwrite = bool(params.get("overwrite", False))
        return self._ctx.run_operation(
            method_name="vba.sync_module",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.sync_module(
                workbook_id=workbook_id,
                module_name=module_name,
                module_type=module_type,
                code=code,
                overwrite=overwrite,
            ),
            args_summary={"workbook_id": workbook_id, "module_name": module_name, "module_type": module_type, "overwrite": overwrite},
            default_workbook_id=workbook_id,
        )

    def remove_module(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        module_name = str(params.get("module_name", ""))
        return self._ctx.run_operation(
            method_name="vba.remove_module",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.remove_module(
                workbook_id=workbook_id,
                module_name=module_name,
            ),
            args_summary={"workbook_id": workbook_id, "module_name": module_name},
            default_workbook_id=workbook_id,
        )

    def execute_macro(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        procedure_name = str(params.get("procedure_name", ""))
        arguments = params.get("arguments") or []
        timeout_seconds = int(params.get("timeout_seconds", self._ctx.services.config.vba_policy.execution_timeout_seconds))
        return self._ctx.run_operation(
            method_name="vba.execute_macro",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.execute_macro(
                workbook_id=workbook_id,
                procedure_name=procedure_name,
                arguments=arguments,
                timeout_seconds=timeout_seconds,
            ),
            args_summary={"workbook_id": workbook_id, "procedure_name": procedure_name},
            default_workbook_id=workbook_id,
        )

    def execute_inline(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        code = str(params.get("code", ""))
        procedure_name = str(params.get("procedure_name", "Main"))
        timeout_seconds = int(params.get("timeout_seconds", self._ctx.services.config.vba_policy.execution_timeout_seconds))
        return self._ctx.run_operation(
            method_name="vba.execute_inline",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.execute_inline(
                workbook_id=workbook_id,
                code=code,
                procedure_name=procedure_name,
                timeout_seconds=timeout_seconds,
            ),
            args_summary={"workbook_id": workbook_id, "procedure_name": procedure_name},
            default_workbook_id=workbook_id,
        )

    def compile(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        return self._ctx.run_operation(
            method_name="vba.compile",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.vba_service.compile_project(workbook_id=workbook_id),
            args_summary={"workbook_id": workbook_id},
            default_workbook_id=workbook_id,
        )
