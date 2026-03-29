from __future__ import annotations

import argparse
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excelforge.gateway.config import GatewayConfig, load_gateway_config
from excelforge.gateway.logging_setup import setup_logging, get_current_log_file
from excelforge.gateway.profile_resolver import BundleRegistry, ProfileResolutionError, ProfileResolver
from excelforge.gateway.runtime_client_manager import get_global_runtime_client
from excelforge.gateway.runtime_identity import (
    RuntimeIdentity,
    resolve_runtime_identity,
)
from excelforge.gateway.utils import call_runtime


# ── 工具 → runtime 方法映射 ──────────────────────────
TOOL_MANIFEST_MAP: dict[str, str] = {
    "server.get_status": "server.status",
    "server.health": "server.health",
    "workbook.open_file": "workbook.open",
    "workbook.create_file": "workbook.create",
    "workbook.save_file": "workbook.save",
    "workbook.close_file": "workbook.close",
    "workbook.inspect": "workbook.info",
    "workbook.list_open": "workbook.list",
    "names.inspect": "names.list",
    "names.manage": "names.read",
    "names.create": "names.create",
    "names.delete": "names.delete",
    "sheet.create_sheet": "sheet.create",
    "sheet.rename_sheet": "sheet.rename",
    "sheet.preview_delete": "sheet.preview_delete",
    "sheet.delete_sheet": "sheet.delete",
    "sheet.inspect_structure": "sheet.inspect",
    "sheet.set_auto_filter": "sheet.auto_filter",
    "sheet.get_conditional_formats": "sheet.get_conditional_formats",
    "sheet.get_data_validations": "sheet.get_data_validations",
    "range.read_values": "range.read",
    "range.write_values": "range.write",
    "range.clear_contents": "range.clear",
    "range.copy": "range.copy",
    "range.insert_rows": "range.insert_rows",
    "range.delete_rows": "range.delete_rows",
    "range.insert_columns": "range.insert_columns",
    "range.delete_columns": "range.delete_columns",
    "range.sort_data": "range.sort",
    "range.merge": "range.merge",
    "range.unmerge": "range.unmerge",
    "format.set_number_format": "format.set_style",
    "format.set_font": "format.set_style",
    "format.set_fill": "format.set_style",
    "format.set_border": "format.set_style",
    "format.set_alignment": "format.set_style",
    "format.set_column_width": "format.auto_fit",
    "format.set_row_height": "format.auto_fit",
    "formula.fill_range": "formula.fill",
    "formula.set_single": "formula.set_single",
    "formula.get_dependencies": "formula.get_dependencies",
    "formula.repair": "formula.repair",
    "vba.inspect_project": "vba.inspect_project",
    "vba.get_module_code": "vba.get_module_code",
    "vba.scan_code": "vba.scan_code",
    "vba.sync_module": "vba.sync_module",
    "vba.remove_module": "vba.remove_module",
    "vba.execute": "vba.execute_macro",
    "vba.import_module": "vba.import_module",
    "vba.export_module": "vba.export_module",
    "vba.compile": "vba.compile",
    "rollback.manage": "recovery.undo_last",
    "rollback.preview_snapshot": "recovery.preview_snapshot",
    "rollback.restore_snapshot": "recovery.restore_snapshot",
    "snapshot.manage": "recovery.list_snapshots",
    "snapshot.get_stats": "recovery.snapshot_stats",
    "snapshot.cleanup": "recovery.snapshot_cleanup",
    "backups.manage": "recovery.list_backups",
    "backups.restore": "recovery.restore_backup",
    "pq.list_connections": "pq.list_connections",
    "pq.list_queries": "pq.list_queries",
    "pq.get_code": "pq.get_query_code",
    "pq.update_query": "pq.update_query",
    "pq.refresh": "pq.refresh",
    "audit.list_operations": "audit.list_operations",
    "table.list_tables": "table.list_tables",
    "table.create": "table.create",
    "table.inspect": "table.inspect",
    "table.resize": "table.resize",
    "table.rename": "table.rename",
    "table.set_style": "table.set_style",
    "table.toggle_total_row": "table.toggle_total_row",
    "table.delete": "table.delete",
    "analysis.scan_structure": "analysis.scan_structure",
    "analysis.scan_formulas": "analysis.scan_formulas",
    "analysis.scan_links": "analysis.scan_links",
    "analysis.scan_hidden": "analysis.scan_hidden",
    "analysis.export_report": "analysis.export_report",
    "workbook.save_as": "workbook.save_as",
    "workbook.refresh_all": "workbook.refresh_all",
    "workbook.calculate": "workbook.calculate",
    "workbook.list_links": "workbook.list_links",
    "workbook.export_pdf": "workbook.export_pdf",
    "sheet.export_csv": "sheet.export_csv",
}

# ── 工具参数 JSON Schema 注册表 ─────────────────────
# 每个 tool_name → (description, {param_name: json_schema})
# 与 runtime_api 中各方法从 params dict 读取的 key 一一对应
_STR = {"type": "string"}
_BOOL = {"type": "boolean"}
_INT = {"type": "integer"}

TOOL_PARAM_SCHEMA: dict[str, tuple[str, dict[str, dict]]] = {
    # ── server ──
    "server.get_status": ("获取 ExcelForge Runtime 状态", {}),
    "server.health": ("健康检查", {}),
    # ── workbook ──
    "workbook.open_file": ("打开 Excel 工作簿", {
        "file_path": {**_STR, "description": "工作簿的绝对路径"},
        "read_only": {**_BOOL, "description": "是否只读打开", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.create_file": ("创建新 Excel 工作簿", {
        "file_path": {**_STR, "description": "新工作簿的保存路径（绝对路径）"},
        "sheet_names": {"type": "array", "items": {"type": "string"}, "description": "工作表名称列表", "default": ["Sheet1"]},
        "overwrite": {**_BOOL, "description": "是否覆盖已存在的文件", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.save_file": ("保存工作簿", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "save_as_path": {"type": "string", "description": "另存为路径（可选，不填则原位保存）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.close_file": ("关闭工作簿", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "save_before_close": {**_BOOL, "description": "关闭前是否保存", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.list_open": ("列出已打开的工作簿", {
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.inspect": ("获取工作簿信息", {
        "workbook_id": {"type": "string", "description": "工作簿 ID（可选，不填返回默认工作簿）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── names ──
    "names.inspect": ("列出命名区域", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "scope": {**_STR, "description": "范围: all / workbook / sheet", "default": "all"},
        "sheet_name": {"type": "string", "description": "工作表名称（scope=sheet 时必填）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "names.manage": ("读取命名区域值", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "range_name": {**_STR, "description": "命名区域名称"},
        "value_mode": {**_STR, "description": "值模式: raw / formatted / formula", "default": "raw"},
        "row_offset": {**_INT, "description": "行偏移", "default": 0},
        "row_limit": {**_INT, "description": "最大行数", "default": 200},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "names.create": ("创建命名区域", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "name": {**_STR, "description": "命名区域名称"},
        "refers_to": {**_STR, "description": "引用区域（如 Sheet1!$A$1:$D$10）"},
        "scope": {**_STR, "description": "作用域: workbook / sheet", "default": "workbook"},
        "sheet_name": {"type": "string", "description": "工作表名称（scope=sheet 时必填）", "default": None},
        "overwrite": {**_BOOL, "description": "是否覆盖已有同名命名区域", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "names.delete": ("删除命名区域", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "name": {**_STR, "description": "命名区域名称"},
        "scope": {**_STR, "description": "作用域: workbook / sheet", "default": "workbook"},
        "sheet_name": {"type": "string", "description": "工作表名称（scope=sheet 时必填）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── sheet ──
    "sheet.create_sheet": ("创建工作表", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "新工作表名称"},
        "position": {"type": "string", "description": "位置: last（默认，末尾）/ first / after（激活表之后）/ before / after", "default": "last"},
        "reference_sheet": {"type": "string", "description": "参考工作表（position=before/after 时使用）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.rename_sheet": ("重命名工作表", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "当前工作表名称"},
        "new_name": {**_STR, "description": "新名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.preview_delete": ("预览删除工作表（获取 confirm_token）", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "要删除的工作表名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.delete_sheet": ("删除工作表", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "confirm_token": {"type": "string", "description": "确认令牌（从预览获取）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.inspect_structure": ("检查工作表结构", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "sample_rows": {**_INT, "description": "预览行数", "default": 5},
        "scan_rows": {**_INT, "description": "扫描行数", "default": 10},
        "max_profile_columns": {**_INT, "description": "最大分析列数", "default": 50},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.set_auto_filter": ("设置自动筛选", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "action": {**_STR, "description": "操作: set / clear"},
        "range": {**_STR, "description": "筛选区域（如 A1:D100）"},
        "filters": {"type": "array", "items": {"type": "object"}, "description": "筛选条件列表"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.get_conditional_formats": ("获取条件格式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {"type": "string", "description": "单元格区域（可选）", "default": None},
        "limit": {**_INT, "description": "最大返回数量", "default": 100},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.get_data_validations": ("获取数据验证规则", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {"type": "string", "description": "单元格区域（可选）", "default": None},
        "limit": {**_INT, "description": "最大返回数量", "default": 100},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── range ──
    "range.read_values": ("读取单元格值", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {**_STR, "description": "单元格区域（如 A1:D10）"},
        "value_mode": {**_STR, "description": "值模式: raw / formatted / formula", "default": "raw"},
        "include_formulas": {**_BOOL, "description": "是否包含公式列", "default": False},
        "row_offset": {**_INT, "description": "行偏移", "default": 0},
        "row_limit": {**_INT, "description": "最大行数", "default": 200},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.write_values": ("写入单元格值", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "start_cell": {**_STR, "description": "起始单元格（如 A1）"},
        "values": {"type": "array", "items": {"type": "object"}, "description": "要写入的值（二维数组）"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.clear_contents": ("清除单元格内容", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {**_STR, "description": "单元格区域"},
        "scope": {**_STR, "description": "清除范围: contents / formats / all", "default": "contents"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.copy": ("复制单元格区域", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "源工作表名称"},
        "source_range": {**_STR, "description": "源区域"},
        "target_sheet": {**_STR, "description": "目标工作表名称"},
        "target_start_cell": {**_STR, "description": "目标起始单元格"},
        "paste_mode": {**_STR, "description": "粘贴模式: values / formulas / formats / all", "default": "values"},
        "target_workbook_id": {"type": "string", "description": "目标工作簿 ID（跨工作簿复制时使用）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.insert_rows": ("插入行", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "row": {**_INT, "description": "在第几行之前插入"},
        "count": {**_INT, "description": "插入行数", "default": 1},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.delete_rows": ("删除行", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "row": {**_INT, "description": "起始行号"},
        "count": {**_INT, "description": "删除行数", "default": 1},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.insert_columns": ("插入列", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "column": {**_STR, "description": "列标识（如 A 或 3）"},
        "count": {**_INT, "description": "插入列数", "default": 1},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.delete_columns": ("删除列", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "column": {**_STR, "description": "列标识（如 A 或 3）"},
        "count": {**_INT, "description": "删除列数", "default": 1},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.sort_data": ("排序数据", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {**_STR, "description": "排序区域"},
        "sort_keys": {"type": "array", "items": {"type": "object"}, "description": "排序键列表（每项含 column/direction）"},
        "has_header": {**_BOOL, "description": "是否有标题行", "default": False},
        "case_sensitive": {**_BOOL, "description": "是否区分大小写", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.merge": ("合并单元格", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {**_STR, "description": "要合并的区域"},
        "across": {**_BOOL, "description": "是否按行合并", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "range.unmerge": ("取消合并单元格", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range": {**_STR, "description": "要取消合并的区域"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── format ──
    "format.set_number_format": ("设置数字格式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "单元格区域"},
        "number_format": {**_STR, "description": "数字格式（如 '0.00', 'yyyy-mm-dd'）"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_font": ("设置字体", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "单元格区域"},
        "name": {**_STR, "description": "字体名称（如 'Arial', '宋体'）（可选）", "default": None},
        "size": {**_INT, "description": "字体大小（可选）", "default": None},
        "bold": {**_BOOL, "description": "是否加粗", "default": False},
        "italic": {**_BOOL, "description": "是否斜体", "default": False},
        "font_color": {**_STR, "description": "字体颜色（如 'FF0000' 红色）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_fill": ("设置填充色", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "单元格区域"},
        "fill_color": {**_STR, "description": "填充颜色（如 'FFFF00' 黄色）"},
        "pattern": {**_STR, "description": "填充图案（如 'solid', 'gray75'）", "default": "solid"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_border": ("设置边框", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "单元格区域"},
        "border_style": {**_STR, "description": "边框样式（如 'thin', 'medium', 'thick'）"},
        "border_type": {**_STR, "description": "边框类型: all/outside/inside/left/right/top/bottom", "default": "all"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_alignment": ("设置对齐方式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "单元格区域"},
        "horizontal": {**_STR, "description": "水平对齐: left/center/right/general", "default": None},
        "vertical": {**_STR, "description": "垂直对齐: top/center/bottom", "default": None},
        "wrap_text": {**_BOOL, "description": "是否自动换行", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_column_width": ("自动调整列宽", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "列区域（如 A:D 或 A）"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "format.set_row_height": ("自动调整行高", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "行区域"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── formula ──
    "formula.fill_range": ("批量填充公式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "目标区域（如 A2:A100）"},
        "formula": {**_STR, "description": "公式表达式（如 =B2*C2）"},
        "formula_type": {**_STR, "description": "公式类型: standard / array / r1c1", "default": "standard"},
        "preview_rows": {**_INT, "description": "预览行数", "default": 5},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "formula.set_single": ("设置单个公式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "cell": {**_STR, "description": "单元格地址（如 A1）"},
        "formula": {**_STR, "description": "公式表达式"},
        "formula_type": {**_STR, "description": "公式类型: standard / array / r1c1", "default": "standard"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "formula.get_dependencies": ("获取公式依赖链", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "cell": {**_STR, "description": "单元格地址（如 A1）"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "formula.repair": ("修复/扫描公式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称（可选，默认激活表）", "default": None},
        "range": {**_STR, "description": "扫描区域"},
        "action": {**_STR, "description": "操作: scan（扫描）/ repair（修复）", "default": "scan"},
        "replacements": {"type": "array", "items": {"type": "object"}, "description": "替换规则列表（action=repair 时使用）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── vba ──
    "vba.inspect_project": ("检查 VBA 工程", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.get_module_code": ("获取 VBA 模块代码", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "module_name": {**_STR, "description": "模块名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.scan_code": ("扫描 VBA 代码风险", {
        "code": {**_STR, "description": "VBA 代码"},
        "module_name": {**_STR, "description": "模块名称"},
        "module_type": {**_STR, "description": "模块类型: standard_module / class_module / userform / document", "default": "standard_module"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.sync_module": ("同步 VBA 模块代码", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "module_name": {**_STR, "description": "模块名称"},
        "module_type": {**_STR, "description": "模块类型: standard_module / class_module / userform / document", "default": "standard_module"},
        "code": {**_STR, "description": "VBA 代码内容"},
        "overwrite": {**_BOOL, "description": "是否覆盖已有代码", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.remove_module": ("删除 VBA 模块", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "module_name": {**_STR, "description": "模块名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.execute": ("执行 VBA 宏", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "procedure_name": {**_STR, "description": "过程名称"},
        "arguments": {"type": "array", "items": {"type": "object"}, "description": "参数列表", "default": []},
        "timeout_seconds": {**_INT, "description": "超时秒数", "default": 30},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.import_module": ("导入 VBA 模块(.bas/.cls)", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "file_path": {**_STR, "description": "要导入的文件路径(.bas或.cls)"},
        "module_name": {**_STR, "description": "模块名称(可选，默认从文件推断)", "default": None},
        "overwrite": {**_BOOL, "description": "如果模块已存在是否覆盖", "default": False},
        "client_request_id": {**_STR, "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.export_module": ("导出 VBA 模块到文件", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "module_name": {**_STR, "description": "要导出的模块名称"},
        "file_path": {**_STR, "description": "导出目标文件路径(.bas或.cls)"},
        "overwrite": {**_BOOL, "description": "如果文件已存在是否覆盖", "default": False},
        "client_request_id": {**_STR, "description": "客户端请求 ID（可选）", "default": None},
    }),
    "vba.compile": ("编译 VBA 工程", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── recovery ──
    "rollback.manage": ("撤销最后一次操作", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "rollback.preview_snapshot": ("预览快照内容", {
        "snapshot_id": {**_STR, "description": "快照 ID"},
        "sample_limit": {**_INT, "description": "预览行数", "default": 20},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "rollback.restore_snapshot": ("恢复指定快照", {
        "snapshot_id": {**_STR, "description": "快照 ID"},
        "preview_token": {"type": "string", "description": "预览令牌（从预览快照获取）", "default": ""},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "snapshot.manage": ("列出快照", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "limit": {**_INT, "description": "最大返回数量", "default": 20},
        "offset": {**_INT, "description": "偏移量", "default": 0},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "snapshot.get_stats": ("获取快照统计信息", {
        "workbook_id": {**_STR, "description": "工作簿 ID（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "snapshot.cleanup": ("清理过期快照", {
        "workbook_id": {**_STR, "description": "工作簿 ID（可选）", "default": None},
        "max_age_hours": {**_INT, "description": "最大保留小时数（可选）", "default": None},
        "dry_run": {**_BOOL, "description": "仅预览不实际删除", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "backups.manage": ("列出备份", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "file_path": {"type": "string", "description": "文件路径筛选", "default": None},
        "limit": {**_INT, "description": "最大返回数量", "default": 20},
        "offset": {**_INT, "description": "偏移量", "default": 0},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "backups.restore": ("恢复指定备份", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "backup_id": {**_STR, "description": "备份 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── pq ──
    "pq.list_connections": ("列出 Power Query 连接", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "pq.list_queries": ("列出 Power Query 查询", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "pq.get_code": ("获取 Power Query 代码", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "query_name": {**_STR, "description": "查询名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "pq.update_query": ("更新 Power Query", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "query_name": {**_STR, "description": "查询名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "pq.refresh": ("刷新 Power Query", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "query_name": {**_STR, "description": "查询名称"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── audit ──
    "audit.list_operations": ("列出操作审计记录", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "tool_name": {"type": "string", "description": "工具名称筛选", "default": None},
        "success_only": {**_BOOL, "description": "仅显示成功操作", "default": False},
        "limit": {**_INT, "description": "最大返回数量", "default": 20},
        "offset": {**_INT, "description": "偏移量", "default": 0},
        "operation_id": {"type": "string", "description": "操作 ID", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── table ──
    "table.list_tables": ("列出工作簿中所有表格", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选，不填则搜索所有表）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.create": ("将区域转换为表格", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "range_address": {**_STR, "description": "要转换为表格的区域（如 A1:D10）"},
        "table_name": {"type": "string", "description": "表格名称（可选，自动生成）", "default": None},
        "has_header": {**_BOOL, "description": "是否有标题行", "default": True},
        "style_name": {"type": "string", "description": "表格样式名称（如 TableStyleMedium2）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.inspect": ("检查表格结构", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "表格名称"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.resize": ("调整表格范围", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "表格名称"},
        "new_range_address": {**_STR, "description": "新的表格区域（如 A1:E20）"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.rename": ("重命名表格", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "当前表格名称"},
        "new_name": {**_STR, "description": "新名称"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.set_style": ("设置表格样式", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "表格名称"},
        "style_name": {"type": "string", "description": "样式名称（可选，不填则移除样式）", "default": None},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.toggle_total_row": ("开关总计行", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "表格名称"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "table.delete": ("删除表格（保留数据）", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "table_name": {**_STR, "description": "表格名称"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── analysis ──
    "analysis.scan_structure": ("扫描工作簿结构", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "analysis.scan_formulas": ("扫描公式分布", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {"type": "string", "description": "工作表名称（可选）", "default": None},
        "scan_range": {"type": "string", "description": "扫描区域（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "analysis.scan_links": ("扫描外部链接", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "analysis.scan_hidden": ("扫描隐藏元素", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "analysis.export_report": ("生成分析报告", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "report_format": {**_STR, "description": "报告格式: text/json", "default": "text"},
        "include_formulas": {**_BOOL, "description": "是否包含公式", "default": False},
        "include_links": {**_BOOL, "description": "是否包含链接", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    # ── workbook_ops ──
    "workbook.save_as": ("另存工作簿", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "save_as_path": {**_STR, "description": "保存路径"},
        "file_format": {"type": "string", "description": "文件格式: xlsx/xlsm/xlsb/csv", "default": None},
        "password": {"type": "string", "description": "文件密码（可选）", "default": None},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.refresh_all": ("刷新所有数据", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.calculate": ("重新计算", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.list_links": ("列出外部链接", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "workbook.export_pdf": ("导出 PDF", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "file_path": {**_STR, "description": "PDF 保存路径"},
        "include_hidden_sheets": {**_BOOL, "description": "是否包含隐藏工作表", "default": False},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
    "sheet.export_csv": ("导出 CSV", {
        "workbook_id": {**_STR, "description": "工作簿 ID"},
        "sheet_name": {**_STR, "description": "工作表名称"},
        "file_path": {**_STR, "description": "CSV 保存路径"},
        "delimiter": {"type": "string", "description": "分隔符", "default": ","},
        "include_header": {**_BOOL, "description": "是否包含表头", "default": True},
        "client_request_id": {"type": "string", "description": "客户端请求 ID（可选）", "default": None},
    }),
}


@dataclass(frozen=True)
class HostRuntimeSettings:
    identity: RuntimeIdentity
    auto_start: bool
    connect_timeout: int
    call_timeout: int
    runtime_config_path: str | None
    display_name: str


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-mcp",
        description="ExcelForge Unified MCP Host",
    )
    parser.add_argument(
        "--config",
        help="Path to excel-mcp.yaml (optional, uses runtime-config.yaml by default)",
    )
    parser.add_argument(
        "--profile",
        default="basic_edit",
        help="Profile name (default: basic_edit)",
    )
    parser.add_argument(
        "--enable-bundle",
        action="append",
        default=[],
        dest="enabled_bundles",
        help="Extra bundles to enable (can be repeated)",
    )
    parser.add_argument(
        "--disable-bundle",
        action="append",
        default=[],
        dest="disabled_bundles",
        help="Bundles to disable (can be repeated)",
    )
    parser.add_argument(
        "--strict-profile",
        action="store_true",
        help="Fail immediately if profile not found",
    )
    parser.add_argument(
        "--restart-runtime",
        choices=["always", "if-stale", "never"],
        default="if-stale",
        help="Runtime restart strategy: always=每次启动都重启; if-stale=超时则重启; never=从不重启（生产）",
    )
    parser.add_argument(
        "--runtime-stale-seconds",
        type=int,
        default=300,
        help="if-stale 模式下，Runtime 多久没有健康更新算过期（默认 300 秒）",
    )
    parser.add_argument(
        "--list-profiles",
        action="store_true",
        help="List available profiles and exit",
    )
    parser.add_argument(
        "--list-bundles",
        action="store_true",
        help="List available bundles and exit",
    )
    parser.add_argument(
        "--runtime-scope",
        default="default",
        help="Runtime scope (default: default)",
    )
    parser.add_argument(
        "--runtime-instance",
        default="default",
        help="Runtime instance name (default: default)",
    )
    parser.add_argument(
        "--print-runtime-endpoint",
        action="store_true",
        help="Print resolved Runtime endpoint on startup",
    )
    parser.add_argument(
        "--dump-tools",
        action="store_true",
        help="Output final tool list for current profile and exit",
    )
    parser.add_argument(
        "--dump-tools-with-index",
        action="store_true",
        help="Output tool list with order index and exit",
    )
    parser.add_argument(
        "--dump-profile-resolution",
        action="store_true",
        help="Output profile/bundle resolution process and exit",
    )
    return parser


def _ensure_runtime_fresh(args, settings: HostRuntimeSettings) -> None:
    """
    确保 Runtime 是最新的。根据 --restart-runtime 策略决定是否重启。

    always:   无条件杀掉旧 Runtime，重新启动
    if-stale: 检查 last_health_ping，超时则重启
    never:    什么都不做（生产模式）
    """
    import logging
    import time
    import datetime
    logger = logging.getLogger(__name__)

    strategy = args.restart_runtime
    logger.info(f"[Startup] Runtime restart strategy: {strategy}")

    if strategy == "never":
        logger.info("[Startup] Skipping Runtime restart check")
        return

    runtime_mgr = get_global_runtime_client(
        identity=settings.identity,
        auto_start=settings.auto_start,
        connect_timeout=settings.connect_timeout,
        call_timeout=settings.call_timeout,
        runtime_config_path=settings.runtime_config_path,
    )

    if strategy == "always":
        logger.info("[Startup] Forcing Runtime restart (--restart-runtime=always)")
        _kill_and_restart_runtime(runtime_mgr, settings)
        return

    if strategy == "if-stale":
        stale_seconds = args.runtime_stale_seconds
        try:
            status = runtime_mgr.call("server.status", {})
            last_ping = status.get("meta", {}).get("last_health_ping") if isinstance(status, dict) else None
            if last_ping:
                if isinstance(last_ping, str):
                    ping_time = datetime.datetime.fromisoformat(last_ping.replace("Z", "+00:00"))
                    age = (datetime.datetime.now(datetime.timezone.utc) - ping_time).total_seconds()
                else:
                    age = 0
                logger.info(f"[Startup] Runtime last ping: {age:.0f}s ago")
                if age > stale_seconds:
                    logger.warning(f"[Startup] Runtime is stale (>{stale_seconds}s), restarting")
                    _kill_and_restart_runtime(runtime_mgr, settings)
                else:
                    logger.info("[Startup] Runtime is fresh, reusing")
            else:
                logger.info("[Startup] No Runtime last ping, starting fresh")
                _kill_and_restart_runtime(runtime_mgr, settings)
        except Exception as e:
            logger.warning(f"[Startup] Cannot check Runtime status: {e}, starting fresh")
            try:
                _kill_and_restart_runtime(runtime_mgr, settings)
            except Exception:
                pass


def _kill_and_restart_runtime(runtime_mgr, settings: HostRuntimeSettings) -> None:
    """终止旧 Runtime 并启动新的。"""
    import logging
    import time
    logger = logging.getLogger(__name__)

    logger.info("[Startup] Killing old Runtime...")
    try:
        runtime_mgr.kill_runtime()
    except Exception as e:
        logger.warning(f"[Startup] Failed to kill old Runtime: {e}")

    time.sleep(1)

    logger.info("[Startup] Starting new Runtime...")
    try:
        runtime_mgr._start_runtime_process()
        _wait_for_runtime_ready(runtime_mgr, timeout=30)
        logger.info("[Startup] New Runtime is ready")
    except Exception as e:
        logger.error(f"[Startup] Failed to start Runtime: {e}")
        raise


def _wait_for_runtime_ready(runtime_mgr, timeout: int = 30):
    """等待新 Runtime 就绪。"""
    import time
    start = time.time()
    while time.time() - start < timeout:
        try:
            status = runtime_mgr.call("server.status", {})
            if status:
                return
        except Exception:
            pass
        time.sleep(1)
    raise TimeoutError(f"Runtime not ready after {timeout}s")


def list_profiles_and_exit(profiles_path: Path | None = None) -> None:
    resolver = ProfileResolver(profiles_path)
    profiles = resolver.list_profiles()
    print("Available profiles:")
    for name in profiles:
        info = resolver.get_profile_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    bundles: {', '.join(info['bundles'])}")


def list_bundles_and_exit(bundles_path: Path | None = None) -> None:
    registry = BundleRegistry(bundles_path)
    bundles = registry.list_bundles()
    print("Available bundles:")
    for name in bundles:
        info = registry.get_bundle_info(name)
        print(f"  {name}")
        if info["description"]:
            print(f"    {info['description']}")
        print(f"    domains: {', '.join(info['domains'])}")


def dump_tools_for_profile(
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> list[str]:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)
    return enabled_tools


def dump_tools_and_exit(
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)

    print(f"Profile: {profile_name}")
    print(f"Category: {profile_info.get('category', 'unknown')}")
    budget = profile_info.get("tool_budget")
    print(f"Tool Budget: {budget if budget is not None else 'unlimited'}")
    print(f"Actual Tools: {len(enabled_tools)}")
    budget_status = "OK" if budget is None or len(enabled_tools) <= budget else "WARNING"
    print(f"Budget Status: {'OK' if budget_status == 'OK' else 'WARNING'}")
    print()
    print("Bundles:")
    for idx, bundle_name in enumerate(resolved_bundles, start=1):
        bundle_tools = bundle_registry.get_bundle_tools(bundle_name)
        tool_count = len(bundle_tools)
        print(f"  [{idx}] {bundle_name} ({tool_count} tools)")
    print()
    print("Tools:")
    for idx, tool_name in enumerate(enabled_tools, start=1):
        print(f"  {tool_name}")


def dump_tools_with_index_and_exit(
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)

    print(f"Profile: {profile_name}")
    print(f"Category: {profile_info.get('category', 'unknown')}")
    budget = profile_info.get("tool_budget")
    print(f"Tool Budget: {budget if budget is not None else 'unlimited'}")
    print(f"Actual Tools: {len(enabled_tools)}")
    budget_status = "OK" if budget is None or len(enabled_tools) <= budget else "WARNING"
    print(f"Budget Status: {budget_status}")
    print()
    print("Bundles:")
    for idx, bundle_name in enumerate(resolved_bundles, start=1):
        bundle_tools = bundle_registry.get_bundle_tools(bundle_name)
        tool_count = len(bundle_tools)
        print(f"  [{idx}] {bundle_name} ({tool_count} tools)")
    print()
    print("Tools:")
    for idx, tool_name in enumerate(enabled_tools, start=1):
        schema_entry = TOOL_PARAM_SCHEMA.get(tool_name)
        desc = schema_entry[0] if schema_entry else ""
        print(f"  [{idx:02d}] {tool_name}")
        if desc:
            print(f"      {desc}")


def dump_profile_resolution_and_exit(
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    print("=== Profile Resolution ===")
    print(f"Requested profile: {profile_name}")
    print()

    profile_info = resolver.resolve(profile_name)
    print(f"Profile '{profile_name}':")
    print(f"  description: {profile_info.get('description', 'N/A')}")
    print(f"  category: {profile_info.get('category', 'unknown')}")
    print(f"  tool_budget: {profile_info.get('tool_budget', 'unlimited')}")
    print(f"  base bundles: {profile_info.get('bundles', [])}")
    print()

    all_bundles = list(profile_info["bundles"])
    print(f"Extra bundles (--enable-bundle): {extra_bundles}")
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
            print(f"  + {b} added")
    print()

    print(f"Disabled bundles (--disable-bundle): {disabled_bundles}")
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)
            print(f"  - {b} removed")
    print()

    print(f"Final bundle list: {all_bundles}")
    print()

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    print(f"Resolved bundles (with dependencies): {resolved_bundles}")
    print()

    for bundle_name in resolved_bundles:
        bundle_info = bundle_registry.get_bundle_info(bundle_name)
        deps = bundle_info.get("dependencies", [])
        bundle_tools = bundle_registry.get_bundle_tools(bundle_name)
        print(f"Bundle '{bundle_name}':")
        print(f"  description: {bundle_info.get('description', 'N/A')}")
        print(f"  dependencies: {deps if deps else 'none'}")
        print(f"  tools ({len(bundle_tools)}):")
        for tool in bundle_tools:
            print(f"    - {tool}")
        print()

    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)
    print(f"=== Total: {len(enabled_tools)} tools ===")


def check_tool_budget(tool_count: int, profile_info: dict[str, Any]) -> None:
    budget = profile_info.get("tool_budget")
    if budget is None:
        return
    if tool_count > budget:
        warning_msg = (
            f"[ExcelForge WARNING] Tool count ({tool_count}) exceeds budget ({budget}) "
            f"for profile '{profile_info['name']}'. Consider reducing enabled bundles."
        )
        print(warning_msg, file=sys.stderr)
        warnings.warn(
            f"Tool count ({tool_count}) exceeds budget ({budget}) for profile '{profile_info['name']}'. "
            f"Consider reducing enabled bundles.",
            UserWarning,
        )


def log_profile_summary(profile_info: dict[str, Any], enabled_tools: list[str], resolved_bundles: list[str]) -> None:
    import logging
    logger = logging.getLogger(__name__)

    logger.info("=== Profile Summary ===")
    logger.info("Profile: %s", profile_info.get("name"))
    logger.info("Category: %s", profile_info.get("category", "unknown"))
    logger.info("Description: %s", profile_info.get("description", "N/A"))
    budget = profile_info.get("tool_budget")
    logger.info("Tool Budget: %s", budget if budget is not None else "unlimited")
    logger.info("Resolved Bundles (%d): %s", len(resolved_bundles), resolved_bundles)
    logger.info("Enabled Tools: %d", len(enabled_tools))
    if budget is not None and len(enabled_tools) > budget:
        logger.warning("Tool count EXCEEDS budget! Budget=%d, Actual=%d", budget, len(enabled_tools))
    logger.info("========================")


def _resolve_path(base_dir: Path, raw_path: str | None) -> str | None:
    if not raw_path:
        return None
    path = Path(raw_path)
    if not path.is_absolute():
        path = (base_dir / path).resolve()
    else:
        path = path.resolve()
    return str(path)


def _resolve_gateway_config_path(raw_path: str | None) -> Path | None:
    if raw_path:
        return Path(raw_path).resolve()

    default_path = Path("excel-mcp.yaml")
    if default_path.exists():
        return default_path.resolve()
    return None


def resolve_host_runtime_settings(args: argparse.Namespace) -> HostRuntimeSettings:
    config_path = _resolve_gateway_config_path(args.config)
    gateway_config: GatewayConfig | None = None
    runtime_data_dir: str | None = None
    runtime_config_path = str(Path("runtime-config.yaml").resolve())
    auto_start = True
    connect_timeout = 10
    call_timeout = 30
    display_name = "ExcelForge"

    if config_path is not None:
        gateway_config = load_gateway_config(config_path)
        base_dir = config_path.parent
        runtime_data_dir = _resolve_path(base_dir, gateway_config.gateway.runtime_data_dir)
        runtime_config_path = _resolve_path(base_dir, gateway_config.gateway.runtime_config_path)
        auto_start = gateway_config.gateway.auto_start_runtime
        connect_timeout = gateway_config.gateway.connect_timeout_seconds
        call_timeout = gateway_config.gateway.call_timeout_seconds
        display_name = gateway_config.gateway.display_name

    identity = resolve_runtime_identity(
        runtime_data_dir=runtime_data_dir,
        scope=args.runtime_scope,
        instance_name=args.runtime_instance,
    )
    return HostRuntimeSettings(
        identity=identity,
        auto_start=auto_start,
        connect_timeout=connect_timeout,
        call_timeout=call_timeout,
        runtime_config_path=runtime_config_path,
        display_name=display_name,
    )


def create_host_runtime_client(settings: HostRuntimeSettings) -> Any:
    client = get_global_runtime_client(
        identity=settings.identity,
        auto_start=settings.auto_start,
        connect_timeout=settings.connect_timeout,
        call_timeout=settings.call_timeout,
        runtime_config_path=settings.runtime_config_path,
    )
    return client


def _build_typed_handler(
    runtime_client: Any,
    tool_name: str,
    runtime_method: str,
    param_defs: dict[str, dict],
):
    """
    动态生成带类型注解的 handler 函数。

    FastMCP 根据函数签名中的类型注解自动推断 inputSchema。
    如果用 def handler(**kwargs)，FastMCP 只会看到一个 kwargs 参数，
    导致 MCP 协议层无法暴露具体参数名，WorkBuddy 桥接也无法正确传参。

    通过 exec 动态构建函数签名，确保每个工具都有明确的参数类型声明。
    """
    if not param_defs:
        # 无参数工具（如 server.health）
        def handler() -> dict:
            return call_runtime(runtime_client, tool_name=tool_name, method=runtime_method, params={})
        return handler

    # JSON Schema type → Python 类型
    _type_map = {
        "string": "str",
        "boolean": "bool",
        "integer": "int",
        "number": "float",
        "object": "dict",
        "array": "list",
    }

    # 将参数分为必填和可选，确保必填在前（Python 语法要求）
    required_params = [(n, d) for n, d in param_defs.items() if "default" not in d]
    optional_params = [(n, d) for n, d in param_defs.items() if "default" in d]

    params_code = []
    for pname, pdef in required_params + optional_params:
        py_type = _type_map.get(pdef.get("type", "string"), "str")
        if "default" in pdef:
            default_val = pdef["default"]
            if isinstance(default_val, str):
                default_repr = f'"{default_val}"'
            elif isinstance(default_val, bool):
                default_repr = "True" if default_val else "False"
            elif default_val is None:
                default_repr = "None"
            else:
                default_repr = repr(default_val)
            params_code.append(f"{pname}: {py_type} = {default_repr}")
        else:
            params_code.append(f"{pname}: {py_type}")

    params_str = ", ".join(params_code)
    result_items = ", ".join(f'"{p}": {p}' for p in param_defs.keys())

    # 通过 exec 构建带正确签名和注解的函数
    func_code = f"""def _handler({params_str}) -> dict:
    return {{"_tool": "{tool_name}", "_params": {{{result_items}}}}}"""
    ns: dict[str, Any] = {"__builtins__": __builtins__}
    exec(func_code, ns)

    # 取出动态函数，绑定 runtime_client 和 runtime_method 为闭包变量
    _inner = ns["_handler"]
    client_ref = runtime_client
    method_ref = runtime_method
    tool_ref = tool_name

    def handler(**kwargs) -> dict:
        return call_runtime(client_ref, tool_name=tool_ref, method=method_ref, params=kwargs)

    # 复制动态函数的签名和注解到实际 handler，让 FastMCP 能正确推断 inputSchema
    import inspect as _inspect
    handler.__signature__ = _inspect.signature(_inner)  # type: ignore[attr-defined]
    handler.__annotations__ = _inner.__annotations__  # type: ignore[attr-defined]
    handler.__defaults__ = _inner.__defaults__  # type: ignore[attr-defined]
    handler.__kwdefaults__ = _inner.__kwdefaults__  # type: ignore[attr-defined]

    return handler


def register_tools_for_profile(
    mcp: FastMCP,
    runtime: Any,
    profile_name: str,
    extra_bundles: list[str],
    disabled_bundles: list[str],
    profiles_path: Path | None = None,
    bundles_path: Path | None = None,
) -> None:
    resolver = ProfileResolver(profiles_path)
    bundle_registry = BundleRegistry(bundles_path)

    profile_info = resolver.resolve(profile_name)
    all_bundles = list(profile_info["bundles"])
    for b in extra_bundles:
        if b not in all_bundles:
            all_bundles.append(b)
    for b in disabled_bundles:
        if b in all_bundles:
            all_bundles.remove(b)

    resolved_bundles = bundle_registry.resolve_bundles(all_bundles)
    enabled_tools = bundle_registry.get_all_tools(resolved_bundles)

    import logging
    logger = logging.getLogger(__name__)
    logger.info("[Host] Profile=%s", profile_name)
    logger.info("[Host] Bundles=%s", all_bundles)
    logger.info("[Host] Resolved bundles=%s", resolved_bundles)
    logger.info("[Host] Resolved tools (%d)=%s", len(enabled_tools), enabled_tools)
    logger.info("[Host] Resolved VBA tools=%s", [t for t in enabled_tools if t.startswith("vba.")])

    check_tool_budget(len(enabled_tools), profile_info)
    log_profile_summary(profile_info, enabled_tools, resolved_bundles)

    registered_vba_tools = []
    for tool_name in enabled_tools:
        runtime_method = TOOL_MANIFEST_MAP.get(tool_name, tool_name)
        schema_entry = TOOL_PARAM_SCHEMA.get(tool_name)
        desc = schema_entry[0] if schema_entry else tool_name
        param_defs = schema_entry[1] if schema_entry else {}

        # 动态生成带类型注解的 handler，让 FastMCP 从函数签名推断正确的 inputSchema
        handler = _build_typed_handler(runtime, tool_name, runtime_method, param_defs)

        mcp.tool(
            name=tool_name,
            description=desc,
            annotations=ToolAnnotations(readOnlyHint=False),
        )(handler)
        if tool_name.startswith("vba."):
            registered_vba_tools.append(tool_name)

    logger.info("[Host] Actually registered VBA tools=%s", registered_vba_tools)


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    profiles_path = Path(__file__).parent / "profiles.yaml"
    bundles_path = Path(__file__).parent / "bundles.yaml"

    log_file = setup_logging()
    print(f"[ExcelForge] Log file: {log_file}", file=sys.stderr)

    if args.list_profiles:
        list_profiles_and_exit(profiles_path)
        return 0

    if args.list_bundles:
        list_bundles_and_exit(bundles_path)
        return 0

    if args.dump_profile_resolution:
        dump_profile_resolution_and_exit(
            profile_name=args.profile,
            extra_bundles=args.enabled_bundles,
            disabled_bundles=args.disabled_bundles,
            profiles_path=profiles_path,
            bundles_path=bundles_path,
        )
        return 0

    if args.dump_tools:
        dump_tools_and_exit(
            profile_name=args.profile,
            extra_bundles=args.enabled_bundles,
            disabled_bundles=args.disabled_bundles,
            profiles_path=profiles_path,
            bundles_path=bundles_path,
        )
        return 0

    if args.dump_tools_with_index:
        dump_tools_with_index_and_exit(
            profile_name=args.profile,
            extra_bundles=args.enabled_bundles,
            disabled_bundles=args.disabled_bundles,
            profiles_path=profiles_path,
            bundles_path=bundles_path,
        )
        return 0

    try:
        settings = resolve_host_runtime_settings(args)
    except Exception as exc:
        print(f"Error resolving settings: {exc}", file=sys.stderr)
        return 1

    _ensure_runtime_fresh(args, settings)

    try:
        runtime = create_host_runtime_client(settings)
    except Exception as exc:
        print(f"Error creating Runtime client: {exc}", file=sys.stderr)
        return 1

    if args.print_runtime_endpoint:
        print(f"Runtime endpoint: {settings.identity.pipe_name}")
        print(f"Runtime instance ID: {settings.identity.instance_id}")

    display_name = f"{settings.display_name} ({args.profile})"
    mcp = FastMCP(display_name)

    register_tools_for_profile(
        mcp=mcp,
        runtime=runtime,
        profile_name=args.profile,
        extra_bundles=args.enabled_bundles,
        disabled_bundles=args.disabled_bundles,
        profiles_path=profiles_path,
        bundles_path=bundles_path,
    )

    try:
        mcp.run(transport="stdio")
    finally:
        runtime.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
