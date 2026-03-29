from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from excelforge.models.vba_models import (
    VbaCompileRequest,
    VbaExecuteMacroRequest,
    VbaGetModuleCodeRequest,
    VbaImportModuleRequest,
    VbaInspectProjectRequest,
    VbaScanCodeRequest,
    VbaSyncModuleRequest,
    VbaRemoveModuleRequest,
    VbaExportModuleRequest,
)
from excelforge.tools.registry import ToolRegistry


def register_vba_tools(mcp: FastMCP, ctx: Any, registry: ToolRegistry) -> None:
    @mcp.tool(name="vba.inspect_project")
    def vba_inspect_project(workbook_id: str, client_request_id: str = "") -> dict:
        req = VbaInspectProjectRequest(workbook_id=workbook_id, client_request_id=client_request_id)
        envelope = ctx.operation_service.run(
            tool_name="vba.inspect_project",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.inspect_project(workbook_id=req.workbook_id),
            args_summary={"workbook_id": req.workbook_id},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.inspect_project", "vba_tools", "vba")

    @mcp.tool(name="vba.get_module_code")
    def vba_get_module_code(
        workbook_id: str,
        module_name: str,
        client_request_id: str = "",
    ) -> dict:
        req = VbaGetModuleCodeRequest(
            workbook_id=workbook_id,
            module_name=module_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.get_module_code",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.get_module_code(
                workbook_id=req.workbook_id,
                module_name=req.module_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "module_name": req.module_name},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.get_module_code", "vba_tools", "vba")

    @mcp.tool(name="vba.scan_code")
    def vba_scan_code(
        code: str,
        module_name: str = "",
        module_type: str = "standard_module",
        client_request_id: str = "",
    ) -> dict:
        req = VbaScanCodeRequest(
            code=code,
            module_name=module_name,
            module_type=module_type,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.scan_code",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.scan_code(
                code=req.code,
                module_name=req.module_name,
                module_type=req.module_type,
            ),
            args_summary={"code_length": len(req.code), "module_name": req.module_name},
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.scan_code", "vba_tools", "vba")

    @mcp.tool(name="vba.sync_module")
    def vba_sync_module(
        workbook_id: str,
        module_name: str = "Module1",
        code: str = "",
        module_type: str = "standard_module",
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = VbaSyncModuleRequest(
            workbook_id=workbook_id,
            module_name=module_name,
            code=code,
            module_type=module_type,
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.sync_module",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.sync_module(
                workbook_id=req.workbook_id,
                module_name=req.module_name,
                module_type=req.module_type,
                code=req.code,
                overwrite=req.overwrite,
            ),
            args_summary={"workbook_id": req.workbook_id, "module_name": req.module_name},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.sync_module", "vba_tools", "vba")

    @mcp.tool(name="vba.remove_module")
    def vba_remove_module(
        workbook_id: str,
        module_name: str = "",
        client_request_id: str = "",
    ) -> dict:
        req = VbaRemoveModuleRequest(
            workbook_id=workbook_id,
            module_name=module_name,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.remove_module",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.remove_module(
                workbook_id=req.workbook_id,
                module_name=req.module_name,
            ),
            args_summary={"workbook_id": req.workbook_id, "module_name": req.module_name},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.remove_module", "vba_tools", "vba")

    @mcp.tool(name="vba.execute")
    def vba_execute(
        workbook_id: str,
        action: str,
        procedure_name: str = "",
        code: str = "",
        arguments: list | None = None,
        timeout_seconds: int = 30,
        client_request_id: str = "",
    ) -> dict:
        if action == "macro":
            req = VbaExecuteMacroRequest(
                workbook_id=workbook_id,
                procedure_name=procedure_name or "",
                arguments=arguments or [],
                timeout_seconds=timeout_seconds,
                client_request_id=client_request_id,
            )
            envelope = ctx.operation_service.run(
                tool_name="vba.execute",
                client_request_id=req.client_request_id,
                operation_fn=lambda: ctx.vba_service.execute_macro(
                    workbook_id=req.workbook_id,
                    procedure_name=req.procedure_name,
                    arguments=req.arguments,
                    timeout_seconds=req.timeout_seconds,
                ),
                args_summary={"workbook_id": req.workbook_id, "procedure_name": req.procedure_name, "action": action},
                default_workbook_id=req.workbook_id,
            )
            return envelope.model_dump(mode="json")
        else:
            return {
                "success": False,
                "code": "E400_BAD_REQUEST",
                "message": f"Invalid action: {action}. Must be 'macro'",
            }

    registry.add("vba.execute", "vba_tools", "vba")

    @mcp.tool(name="vba.import_module")
    def vba_import_module(
        workbook_id: str,
        file_path: str,
        module_name: str | None = None,
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = VbaImportModuleRequest(
            workbook_id=workbook_id,
            file_path=file_path,
            module_name=module_name,
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.import_module",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.import_module(
                workbook_id=req.workbook_id,
                file_path=req.file_path,
                module_name=req.module_name,
                overwrite=req.overwrite,
            ),
            args_summary={"workbook_id": req.workbook_id, "file_path": req.file_path, "module_name": req.module_name},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.import_module", "vba_tools", "vba")

    @mcp.tool(name="vba.export_module")
    def vba_export_module(
        workbook_id: str,
        module_name: str,
        file_path: str,
        overwrite: bool = False,
        client_request_id: str = "",
    ) -> dict:
        req = VbaExportModuleRequest(
            workbook_id=workbook_id,
            module_name=module_name,
            file_path=file_path,
            overwrite=overwrite,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.export_module",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.export_module(
                workbook_id=req.workbook_id,
                module_name=req.module_name,
                file_path=req.file_path,
                overwrite=req.overwrite,
            ),
            args_summary={"workbook_id": req.workbook_id, "module_name": req.module_name, "file_path": req.file_path},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.export_module", "vba_tools", "vba")

    @mcp.tool(name="vba.compile")
    def vba_compile(
        workbook_id: str,
        client_request_id: str = "",
    ) -> dict:
        req = VbaCompileRequest(
            workbook_id=workbook_id,
            client_request_id=client_request_id,
        )
        envelope = ctx.operation_service.run(
            tool_name="vba.compile",
            client_request_id=req.client_request_id,
            operation_fn=lambda: ctx.vba_service.compile_project(
                workbook_id=req.workbook_id,
            ),
            args_summary={"workbook_id": req.workbook_id},
            default_workbook_id=req.workbook_id,
        )
        return envelope.model_dump(mode="json")

    registry.add("vba.compile", "vba_tools", "vba")