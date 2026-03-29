from __future__ import annotations

import re
import threading
import time
import uuid
from pathlib import Path
from typing import Any

from excelforge.config import AppConfig
from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.excel_worker import ExcelWorker
from excelforge.services.backup_service import BackupService
from excelforge.services.vba_scanner import VbaSecurityScanner
from excelforge.utils.file_format import is_bas_or_cls, supports_vba
from excelforge.utils.path_guard import normalize_allowed_path

MODULE_TYPE_MAP = {
    1: "standard_module",
    2: "class_module",
    3: "userform",
    100: "document",
}
MODULE_TYPE_CONST = {
    "standard_module": 1,
    "class_module": 2,
}
PROC_DECL_RE = re.compile(
    r"^\s*(?:Public|Private|Friend|Static)?\s*"
    r"(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"([A-Za-z_][A-Za-z0-9_]*)\b",
    re.IGNORECASE,
)
MODULE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,30}$")
VBA_UNSUPPORTED_FORMATS = {"xlsx", "xls"}

BLOCKING_PATTERNS = [
    (r"\bMsgBox\s+", "Debug.Print "),
    (r"\bMsgBox\(", "Debug.Print ("),
    (r"\bInputBox\(", '"" & Chr(0) & "InputBox('),
]


def sanitize_vba_for_automation(code: str) -> str:
    """将阻塞性交互函数替换为安全替代"""
    for pattern, replacement in BLOCKING_PATTERNS:
        code = re.sub(pattern, replacement, code, flags=re.IGNORECASE)
    return code


class VbaService:
    MAX_REBUILD_COUNT = 3
    COOLDOWN_MS = 300

    def __init__(
        self,
        config: AppConfig,
        worker: ExcelWorker,
        backup_service: BackupService,
    ) -> None:
        self._config = config
        self._worker = worker
        self._backup_service = backup_service
        self._scanner = VbaSecurityScanner(
            block_levels=["critical", "high"],
            warn_levels=["medium"],
        )
        self._last_execution_time: dict[str, float] = {}

    def _check_worker_health(self, ctx: Any) -> None:
        worker_state = getattr(self._worker, 'state', 'unknown')
        rebuild_count = getattr(self._worker, 'rebuild_count', 0)
        if worker_state != "running":
            raise ExcelForgeError(
                ErrorCode.E503_EXCEL_WORKER_STOPPED,
                f"Excel Worker is {worker_state}. Please restart the MCP service.",
            )
        if rebuild_count >= self.MAX_REBUILD_COUNT:
            raise ExcelForgeError(
                ErrorCode.E503_EXCEL_REBUILD_LIMIT_REACHED,
                f"Excel Worker rebuild limit reached ({rebuild_count}). Please restart the MCP service.",
            )

    def _check_stale_workbook(self, ctx: Any, workbook_id: str) -> None:
        if ctx.registry.is_stale_workbook_id(workbook_id):
            raise ExcelForgeError(
                ErrorCode.E410_WORKBOOK_STALE,
                "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
            )

    def _cooldown(self, workbook_id: str) -> None:
        last_time = self._last_execution_time.get(workbook_id, 0)
        elapsed = (time.time() - last_time) * 1000
        if elapsed < self.COOLDOWN_MS:
            time.sleep((self.COOLDOWN_MS - elapsed) / 1000)
        self._last_execution_time[workbook_id] = time.time()

    def inspect_project(self, *, workbook_id: str) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                if ctx.registry.is_stale_workbook_id(workbook_id):
                    raise ExcelForgeError(
                        ErrorCode.E410_WORKBOOK_STALE,
                        "Workbook handle is stale after Excel rebuild; reopen with workbook.open_file",
                    )
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            vb_project = self._get_vb_project(handle.workbook_obj)
            protection = "locked" if int(vb_project.Protection) == 1 else "none"
            if protection == "locked":
                return {
                    "workbook_id": workbook_id,
                    "project_name": str(vb_project.Name),
                    "protection": "locked",
                    "modules": [],
                    "references": [],
                    "total_modules": 0,
                    "total_code_lines": 0,
                }

            modules: list[dict[str, Any]] = []
            total_lines = 0
            for component in vb_project.VBComponents:
                code_module = component.CodeModule
                line_count = int(code_module.CountOfLines)
                procedures = self._extract_procedures_with_bounds(code_module)
                modules.append(
                    {
                        "name": str(component.Name),
                        "type": MODULE_TYPE_MAP.get(int(component.Type), "document"),
                        "line_count": line_count,
                        "procedure_count": len(procedures),
                        "procedures": [
                            {
                                "name": p["name"],
                                "kind": p["kind"],
                                "start_line": p["start_line"],
                            }
                            for p in procedures
                        ],
                    }
                )
                total_lines += line_count

            references: list[dict[str, Any]] = []
            for ref in vb_project.References:
                references.append(
                    {
                        "name": str(ref.Name),
                        "description": str(getattr(ref, "Description", "") or ""),
                        "major": int(getattr(ref, "Major", 0) or 0),
                        "minor": int(getattr(ref, "Minor", 0) or 0),
                        "is_broken": bool(getattr(ref, "IsBroken", False)),
                    }
                )

            return {
                "workbook_id": workbook_id,
                "project_name": str(vb_project.Name),
                "protection": "none",
                "modules": modules,
                "references": references,
                "total_modules": len(modules),
                "total_code_lines": total_lines,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def get_module_code(self, *, workbook_id: str, module_name: str) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            vb_project = self._get_vb_project(handle.workbook_obj)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            component = None
            for comp in vb_project.VBComponents:
                if str(comp.Name) == module_name:
                    component = comp
                    break
            if component is None:
                raise ExcelForgeError(ErrorCode.E404_VBA_MODULE_NOT_FOUND, f"VBA module not found: {module_name}")

            code_module = component.CodeModule
            total_lines = int(code_module.CountOfLines)
            code = ""
            if total_lines > 0:
                code = str(code_module.Lines(1, total_lines))

            max_size = int(self._config.limits.max_vba_code_size_bytes)
            truncated = len(code) > max_size
            if truncated:
                code = code[:max_size] + "\n' [Code truncated due to max_vba_code_size_bytes limit]"

            procedures = self._extract_procedures_with_bounds(code_module)
            return {
                "workbook_id": workbook_id,
                "module_name": module_name,
                "module_type": MODULE_TYPE_MAP.get(int(component.Type), "document"),
                "code": code,
                "line_count": total_lines,
                "truncated": truncated,
                "procedures": procedures,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    @staticmethod
    def _normalize_kind(raw: str) -> str:
        lower = raw.strip().lower()
        if lower == "sub":
            return "Sub"
        if lower == "function":
            return "Function"
        if lower == "property get":
            return "Property Get"
        if lower == "property let":
            return "Property Let"
        return "Property Set"

    def _extract_procedures_with_bounds(self, code_module: Any) -> list[dict[str, Any]]:
        line_count = int(code_module.CountOfLines)
        if line_count <= 0:
            return []
        text = str(code_module.Lines(1, line_count))
        starts: list[dict[str, Any]] = []
        for idx, line in enumerate(text.splitlines(), start=1):
            m = PROC_DECL_RE.match(line)
            if not m:
                continue
            starts.append(
                {
                    "name": m.group(2),
                    "kind": self._normalize_kind(m.group(1)),
                    "start_line": idx,
                }
            )
        for i, proc in enumerate(starts):
            if i + 1 < len(starts):
                proc["end_line"] = int(starts[i + 1]["start_line"]) - 1
            else:
                proc["end_line"] = line_count
        return starts

    @staticmethod
    def _get_vb_project(workbook: Any) -> Any:
        try:
            return workbook.VBProject
        except Exception as exc:
            raise ExcelForgeError(
                ErrorCode.E403_VBA_ACCESS_DENIED,
                "Cannot access VBA project. Enable 'Trust access to the VBA project object model' in Excel.",
            ) from exc

    @staticmethod
    def _read_file_with_encoding(file_path: Path) -> str:
        encodings = ["utf-8", "gbk", "gb2312", "gb18030", "latin-1"]
        for enc in encodings:
            try:
                return file_path.read_text(encoding=enc)
            except UnicodeDecodeError:
                continue
        return file_path.read_text(encoding="utf-8", errors="replace")

    @staticmethod
    def _extract_vb_name(code: str) -> str | None:
        match = re.search(r"Attribute\s+VB_Name\s*=\s*[\"'](\w+)[\"']", code, re.IGNORECASE)
        if match:
            return match.group(1)
        match = re.search(r"^Attribute\s+VB_Name\s*=\s*\"([^\"]+)\"", code, re.MULTILINE | re.IGNORECASE)
        if match:
            return match.group(1)
        return None

    def scan_code(
        self,
        *,
        code: str,
        module_name: str | None = None,
        module_type: str = "standard_module",
    ) -> dict[str, Any]:
        result = self._scanner.scan(code, module_type)
        return result.to_dict()

    def sync_module(
        self,
        *,
        workbook_id: str,
        module_name: str,
        module_type: str,
        code: str,
        overwrite: bool,
    ) -> dict[str, Any]:
        if not MODULE_NAME_RE.match(module_name):
            raise ExcelForgeError(
                ErrorCode.E400_VBA_MODULE_NAME_INVALID,
                f"Invalid VBA module name: {module_name}",
            )
        if len(code) > int(self._config.limits.max_vba_code_size_bytes):
            raise ExcelForgeError(
                ErrorCode.E400_VBA_CODE_TOO_LARGE,
                f"VBA code exceeds max size of {self._config.limits.max_vba_code_size_bytes} bytes",
            )
        code = sanitize_vba_for_automation(code)
        scan_result = self._scanner.scan(code, module_type)
        if scan_result.blocked:
            raise ExcelForgeError(
                ErrorCode.E403_VBA_POLICY_BLOCKED,
                f"VBA code blocked due to security policy: {scan_result.risk_level}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            self._cooldown(workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            ext = handle.file_path.lower().split(".")[-1] if handle.file_path else ""
            if ext in VBA_UNSUPPORTED_FORMATS:
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_VBA_UNSUPPORTED,
                    f"Cannot write VBA to format .{ext}. Use .xlsm instead.",
                )

            vb_project = self._get_vb_project(wb)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="vba.sync_module",
                description=f"VBA sync module {module_name}",
            )

            component = None
            for comp in vb_project.VBComponents:
                if str(comp.Name) == module_name:
                    component = comp
                    break

            action = "created"
            if component is not None:
                if not overwrite:
                    raise ExcelForgeError(
                        ErrorCode.E409_VBA_MODULE_EXISTS,
                        f"Module {module_name} already exists. Set overwrite=true to replace.",
                    )
                code_module = component.CodeModule
                if int(code_module.CountOfLines) > 0:
                    code_module.DeleteLines(1, int(code_module.CountOfLines))
                code_module.AddFromString(code)
                action = "updated"
            else:
                type_const = MODULE_TYPE_CONST.get(module_type, 1)
                component = vb_project.VBComponents.Add(type_const)
                component.Name = module_name
                component.CodeModule.AddFromString(code)
                action = "created"

            scan_summary = {
                "passed": scan_result.passed,
                "blocked": scan_result.blocked,
                "risk_level": scan_result.risk_level,
                "findings_count": len(scan_result.findings),
            }

            return {
                "workbook_id": workbook_id,
                "module_name": module_name,
                "module_type": module_type,
                "action": action,
                "line_count": scan_result.line_count,
                "procedure_names": scan_result.procedure_names,
                "scan_result": scan_summary,
                "backup_id": backup_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def remove_module(
        self,
        *,
        workbook_id: str,
        module_name: str,
    ) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            self._cooldown(workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            vb_project = self._get_vb_project(wb)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            component = None
            module_type = "standard_module"
            for comp in vb_project.VBComponents:
                if str(comp.Name) == module_name:
                    component = comp
                    module_type = MODULE_TYPE_MAP.get(int(comp.Type), "document")
                    break

            if component is None:
                raise ExcelForgeError(ErrorCode.E404_VBA_MODULE_NOT_FOUND, f"VBA module not found: {module_name}")

            if module_type not in ("standard_module", "class_module"):
                raise ExcelForgeError(
                    ErrorCode.E409_VBA_MODULE_SYSTEM_PROTECTED,
                    f"Cannot remove system module type: {module_type}",
                )

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="vba.remove_module",
                description=f"VBA remove module {module_name}",
            )

            vb_project.VBComponents.Remove(component)

            return {
                "workbook_id": workbook_id,
                "removed_module": module_name,
                "module_type": module_type,
                "backup_id": backup_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    @staticmethod
    def _build_macro_candidates(wb_name: str, procedure_name: str) -> list[str]:
        pure_name = procedure_name.split(".")[-1] if "." in procedure_name else procedure_name
        if "!" in procedure_name:
            return [procedure_name]
        candidates = [
            pure_name,
            f"'{wb_name}'!{pure_name}",
            f"{wb_name}!{pure_name}",
            f"'{wb_name}'!{procedure_name}",
            f"{procedure_name}",
        ]
        return candidates

    def execute_macro(
        self,
        *,
        workbook_id: str,
        procedure_name: str,
        arguments: list[Any],
        timeout_seconds: int,
    ) -> dict[str, Any]:
        if not self._config.vba_policy.allow_execution:
            raise ExcelForgeError(
                ErrorCode.E409_VBA_EXECUTION_DISABLED,
                "VBA execution is disabled by configuration",
            )

        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            self._cooldown(workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            ext = Path(handle.file_path).suffix.lower() if handle.file_path else ""
            if not supports_vba(ext):
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_VBA_UNSUPPORTED,
                    f"Cannot execute VBA in .{ext} format. Use .xlsm, .xlsb, or .xls.",
                )

            vb_project = self._get_vb_project(wb)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            code = self._get_procedure_code(vb_project, procedure_name)
            if code is None:
                raise ExcelForgeError(
                    ErrorCode.E404_VBA_PROCEDURE_NOT_FOUND,
                    f"VBA procedure not found: {procedure_name}",
                )

            scan_result = self._scanner.scan(code, "standard_module")
            if scan_result.blocked:
                raise ExcelForgeError(
                    ErrorCode.E403_VBA_EXECUTION_BLOCKED,
                    f"VBA execution blocked due to security policy: {scan_result.risk_level}",
                )

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="vba.execute_macro",
                description=f"VBA execute macro {procedure_name}",
            )

            app = ctx.app_manager.ensure_app()
            app.DisplayAlerts = False
            app.EnableEvents = False
            result = None
            executed = False
            execution_time_ms = 0
            timeout_triggered = False

            def mark_unhealthy():
                nonlocal timeout_triggered
                timeout_triggered = True
                ctx.worker.mark_unhealthy("VBA execution timeout")

            timer = threading.Timer(timeout_seconds, mark_unhealthy)
            timer.start()
            try:
                start_time = time.time()
                macro_candidates = self._build_macro_candidates(wb.Name, procedure_name)
                last_exc = None
                for macro_ref in macro_candidates:
                    try:
                        result = app.Run(macro_ref, *arguments)
                        execution_time_ms = int((time.time() - start_time) * 1000)
                        executed = True
                        break
                    except Exception as exc:
                        last_exc = exc
                else:
                    raise last_exc
            except Exception as exc:
                if timeout_triggered:
                    raise ExcelForgeError(
                        ErrorCode.E408_VBA_EXECUTION_TIMEOUT,
                        f"VBA execution timed out after {timeout_seconds} seconds",
                    ) from exc
                raise
            finally:
                timer.cancel()
                try:
                    app.DisplayAlerts = True
                    app.EnableEvents = True
                except Exception:
                    pass

            return_value = self._convert_return_value(result)

            scan_summary = {
                "passed": scan_result.passed,
                "blocked": scan_result.blocked,
                "risk_level": scan_result.risk_level,
                "findings_count": len(scan_result.findings),
            }

            return {
                "workbook_id": workbook_id,
                "procedure_name": procedure_name,
                "executed": executed,
                "return_value": return_value,
                "execution_time_ms": execution_time_ms,
                "scan_result": scan_summary,
                "backup_id": backup_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=timeout_seconds + 10,
            requires_excel=True,
        )

    def export_module(
        self,
        *,
        workbook_id: str,
        module_name: str,
        file_path: str,
        overwrite: bool,
    ) -> dict[str, Any]:
        normalized = normalize_allowed_path(file_path, [Path(p) for p in self._config.paths.allowed_roots], set(self._config.paths.allowed_extensions))
        ext = normalized.suffix.lower()
        if not is_bas_or_cls(ext):
            raise ExcelForgeError(
                ErrorCode.E400_UNSUPPORTED_EXTENSION,
                f"export_module only supports .bas and .cls, not {ext}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            vb_project = self._get_vb_project(handle.workbook_obj)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            component = self._find_component(vb_project, module_name)
            if component is None:
                raise ExcelForgeError(ErrorCode.E404_VBA_MODULE_NOT_FOUND, f"VBA module not found: {module_name}")

            if normalized.exists() and not overwrite:
                raise ExcelForgeError(
                    ErrorCode.E409_FILE_EXISTS,
                    f"Target file already exists: {normalized}. Use overwrite=true to replace.",
                )

            component.Export(str(normalized))
            file_size = normalized.stat().st_size
            line_count = int(component.CodeModule.CountOfLines)

            return {
                "workbook_id": workbook_id,
                "module_name": module_name,
                "module_type": MODULE_TYPE_MAP.get(int(component.Type), "document"),
                "file_path": str(normalized),
                "file_size_bytes": file_size,
                "line_count": line_count,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def import_module(
        self,
        *,
        workbook_id: str,
        file_path: str,
        module_name: str | None,
        overwrite: bool,
    ) -> dict[str, Any]:
        normalized = normalize_allowed_path(file_path, [Path(p) for p in self._config.paths.allowed_roots], set(self._config.paths.allowed_extensions))
        ext = normalized.suffix.lower()
        if not is_bas_or_cls(ext):
            raise ExcelForgeError(
                ErrorCode.E400_UNSUPPORTED_EXTENSION,
                f"import_module only supports .bas and .cls, not {ext}",
            )

        if not normalized.exists():
            raise ExcelForgeError(ErrorCode.E404_BAS_FILE_NOT_FOUND, f"File not found: {normalized}")

        code = self._read_file_with_encoding(normalized)
        detected_name = self._extract_vb_name(code)
        if module_name is None and detected_name is not None:
            module_name = detected_name
        module_type = "standard_module" if ext == ".bas" else "class_module"
        scan_result = self._scanner.scan(code, module_type)
        if scan_result.blocked:
            raise ExcelForgeError(
                ErrorCode.E403_VBA_POLICY_BLOCKED,
                f"VBA code blocked due to security policy: {scan_result.risk_level}",
            )

        def op(ctx: Any) -> dict[str, Any]:
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            ext = Path(handle.file_path).suffix.lower() if handle.file_path else ""
            if not supports_vba(ext):
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_VBA_UNSUPPORTED,
                    f"Cannot import VBA to .{ext} format. Use .xlsm, .xlsb, or .xls.",
                )

            vb_project = self._get_vb_project(wb)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            backup_id, _ = self._backup_service.create_backup(
                workbook=handle,
                source_tool="vba.import_module",
                description=f"VBA import module from {normalized.name}",
            )

            imported = vb_project.VBComponents.Import(str(normalized))
            actual_module_name = str(imported.Name)
            target_name = module_name or actual_module_name

            if target_name != actual_module_name:
                imported.Name = target_name

            existing = self._find_component(vb_project, target_name)
            action = "created"
            if existing is not None and existing != imported:
                if not overwrite:
                    vb_project.VBComponents.Remove(imported)
                    raise ExcelForgeError(
                        ErrorCode.E409_VBA_MODULE_EXISTS,
                        f"Module {target_name} already exists. Use overwrite=true to replace.",
                    )
                vb_project.VBComponents.Remove(existing)
                action = "updated"

            line_count = int(imported.CodeModule.CountOfLines) if imported.CodeModule else 0
            procedures = self._extract_procedure_names(imported.CodeModule) if imported.CodeModule else []

            scan_summary = {
                "passed": scan_result.passed,
                "blocked": scan_result.blocked,
                "risk_level": scan_result.risk_level,
                "findings_count": len(scan_result.findings),
            }

            return {
                "workbook_id": workbook_id,
                "module_name": target_name,
                "module_type": MODULE_TYPE_MAP.get(int(imported.Type), "document"),
                "action": action,
                "line_count": line_count,
                "procedure_names": procedures,
                "scan_result": scan_summary,
                "backup_id": backup_id,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def _find_component(self, vb_project: Any, module_name: str) -> Any | None:
        for comp in vb_project.VBComponents:
            if str(comp.Name) == module_name:
                return comp
        return None

    def compile_project(self, *, workbook_id: str) -> dict[str, Any]:
        def op(ctx: Any) -> dict[str, Any]:
            self._check_worker_health(ctx)
            self._check_stale_workbook(ctx, workbook_id)
            self._cooldown(workbook_id)
            handle = ctx.registry.get(workbook_id)
            if handle is None:
                raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, f"Workbook not open: {workbook_id}")

            wb = handle.workbook_obj
            ext = Path(handle.file_path).suffix.lower() if handle.file_path else ""
            if not supports_vba(ext):
                raise ExcelForgeError(
                    ErrorCode.E409_WORKBOOK_VBA_UNSUPPORTED,
                    f"Cannot compile VBA in .{ext} format. Use .xlsm, .xlsb, or .xls.",
                )

            vb_project = self._get_vb_project(wb)
            if int(vb_project.Protection) == 1:
                raise ExcelForgeError(ErrorCode.E403_VBA_PROJECT_PROTECTED, "VBA project is password protected")

            errors = []
            warnings = []
            method = "unavailable"
            modules_checked = 0
            total_lines = 0

            try:
                app = ctx.app_manager.ensure_app()
                compile_cmd = app.VBE.CommandBars.FindControl(Id=578)
                if compile_cmd and compile_cmd.Enabled:
                    compile_cmd.Execute()
                    method = "vbe_command"
                    return {
                        "workbook_id": workbook_id,
                        "project_name": str(vb_project.Name),
                        "compile_success": True,
                        "method": method,
                        "errors": [],
                        "warnings": [],
                        "modules_checked": 0,
                        "total_lines": 0,
                    }
            except Exception as exc:
                warnings.append(f"VBE CommandBar compile unavailable: {exc}")

            method = "syntax_check"
            for component in vb_project.VBComponents:
                code_module = component.CodeModule
                if int(code_module.CountOfLines) == 0:
                    continue
                modules_checked += 1
                total_lines += int(code_module.CountOfLines)
                module_errors = self._check_module_syntax(component)
                errors.extend(module_errors)

            return {
                "workbook_id": workbook_id,
                "project_name": str(vb_project.Name),
                "compile_success": len(errors) == 0,
                "method": method,
                "errors": errors,
                "warnings": warnings,
                "modules_checked": modules_checked,
                "total_lines": total_lines,
            }

        return self._worker.submit(
            op,
            timeout_seconds=self._config.limits.operation_timeout_seconds,
            requires_excel=True,
        )

    def _check_module_syntax(self, component: Any) -> list[dict[str, Any]]:
        errors = []
        code_module = component.CodeModule
        if int(code_module.CountOfLines) == 0:
            return errors
        text = str(code_module.Lines(1, int(code_module.CountOfLines)))

        subs = len(re.findall(r"^\s*(?:Public|Private|Friend|Static)?\s*Sub\s", text, re.MULTILINE))
        end_subs = len(re.findall(r"^\s*End Sub\s*$", text, re.MULTILINE))
        if subs != end_subs:
            errors.append({
                "module_name": str(component.Name),
                "line_number": None,
                "error_message": f"Sub/End Sub mismatch: {subs} Sub vs {end_subs} End Sub",
                "code_excerpt": None,
            })

        funcs = len(re.findall(r"^\s*(?:Public|Private|Friend|Static)?\s*Function\s", text, re.MULTILINE))
        end_funcs = len(re.findall(r"^\s*End Function\s*$", text, re.MULTILINE))
        if funcs != end_funcs:
            errors.append({
                "module_name": str(component.Name),
                "line_number": None,
                "error_message": f"Function/End Function mismatch: {funcs} Function vs {end_funcs} End Function",
                "code_excerpt": None,
            })

        props = len(re.findall(r"^\s*(?:Public|Private|Friend|Static)?\s*Property\s+(?:Get|Let|Set)\s", text, re.MULTILINE))
        end_props = len(re.findall(r"^\s*End Property\s*$", text, re.MULTILINE))
        if props != end_props:
            errors.append({
                "module_name": str(component.Name),
                "line_number": None,
                "error_message": f"Property/End Property mismatch: {props} Property vs {end_props} End Property",
                "code_excerpt": None,
            })

        return errors

    def _get_procedure_code(self, vb_project: Any, procedure_name: str) -> str | None:
        for comp in vb_project.VBComponents:
            code_module = comp.CodeModule
            line_count = int(code_module.CountOfLines)
            if line_count <= 0:
                continue
            text = str(code_module.Lines(1, line_count))
            lines = text.splitlines()
            for idx, line in enumerate(lines, start=1):
                m = PROC_DECL_RE.match(line)
                if m and m.group(2).lower() == procedure_name.lower():
                    start = idx
                    end = line_count
                    for j in range(idx + 1, line_count + 1):
                        if j - 1 < len(lines):
                            next_m = PROC_DECL_RE.match(lines[j - 1])
                            if next_m:
                                end = j - 1
                                break
                    return "\n".join(lines[start - 1:end])
        return None

    def _extract_procedure_names(self, code_module: Any) -> list[str]:
        line_count = int(code_module.CountOfLines)
        if line_count <= 0:
            return []
        text = str(code_module.Lines(1, line_count))
        procedures = []
        for line in text.splitlines():
            m = PROC_DECL_RE.match(line)
            if m:
                procedures.append(m.group(2))
        return procedures

    @staticmethod
    def _convert_return_value(value: Any) -> str | int | float | bool | None:
        if value is None:
            return None
        if isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, (list, dict, tuple)):
            return None
        return None
