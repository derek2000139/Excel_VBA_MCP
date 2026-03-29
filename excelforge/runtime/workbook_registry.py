from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.handle_ownership import ensure_workbook_id_owned, is_foreign_workbook_id
from excelforge.utils.ids import parse_workbook_generation

logger = logging.getLogger(__name__)


class WorkbookHandleStaleError(ExcelForgeError):
    """工作簿句柄已失效，需要重新打开。"""
    def __init__(self, workbook_id: str):
        super().__init__(
            ErrorCode.E410_WORKBOOK_STALE,
            f"Workbook handle is stale (workbook_id={workbook_id}). Please reopen the workbook."
        )
        self.workbook_id = workbook_id


@dataclass
class WorkbookHandle:
    workbook_id: str
    workbook_name: str
    file_path: str
    read_only: bool
    opened_at: str
    workbook_obj: Any
    file_format: str = "xlsx"
    max_rows: int = 1048576
    max_columns: int = 16384


class WorkbookRegistry:
    def __init__(self, runtime_fingerprint: str | None = None) -> None:
        self._items: dict[str, WorkbookHandle] = {}
        self._generation = 1
        self._runtime_fingerprint = runtime_fingerprint

    @property
    def generation(self) -> int:
        return self._generation

    @property
    def runtime_fingerprint(self) -> str | None:
        return self._runtime_fingerprint

    def count(self) -> int:
        return len(self._items)

    def add(self, handle: WorkbookHandle) -> None:
        self._items[handle.workbook_id] = handle

    def get(self, workbook_id: str) -> WorkbookHandle | None:
        ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint)
        if not self._is_current_generation(workbook_id):
            return None
        return self._items.get(workbook_id)

    def require(self, workbook_id: str) -> WorkbookHandle:
        handle = self.get(workbook_id)
        if handle is None:
            raise KeyError(workbook_id)
        return handle

    def remove(self, workbook_id: str) -> WorkbookHandle | None:
        ensure_workbook_id_owned(workbook_id, self._runtime_fingerprint)
        if not self._is_current_generation(workbook_id):
            return None
        return self._items.pop(workbook_id, None)

    def list_items(self) -> list[WorkbookHandle]:
        return list(self._items.values())

    def find_by_path(self, file_path: str) -> WorkbookHandle | None:
        normalized = str(Path(file_path).resolve()).lower()
        for item in self._items.values():
            if str(Path(item.file_path).resolve()).lower() == normalized:
                return item
        return None

    def invalidate_all(self) -> None:
        self._items.clear()

    def bump_generation(self) -> int:
        self._items.clear()
        self._generation += 1
        return self._generation

    def _is_current_generation(self, workbook_id: str) -> bool:
        generation = parse_workbook_generation(workbook_id)
        if generation is None:
            return True
        return generation == self._generation

    def is_foreign_workbook_id(self, workbook_id: str) -> bool:
        return is_foreign_workbook_id(workbook_id, self._runtime_fingerprint)

    def is_stale_workbook_id(self, workbook_id: str) -> bool:
        if self.is_foreign_workbook_id(workbook_id):
            return False
        generation = parse_workbook_generation(workbook_id)
        return generation is not None and generation != self._generation

    def clear_all(self) -> int:
        """清空所有注册的工作簿句柄。返回被清除的数量。"""
        count = len(self._items)
        self._items.clear()
        logger.info(f"[Registry] Cleared all {count} handle(s)")
        return count

    def validate_handle(self, workbook_id: str) -> bool:
        """
        通过访问 COM 对象验证句柄是否有效。

        Returns:
            True 如果句柄有效，False 如果失效
        """
        handle = self._items.get(workbook_id)
        if handle is None:
            return False
        try:
            _ = handle.workbook_obj.Name
            return True
        except Exception:
            logger.warning(f"[Registry] Stale handle detected: {workbook_id}")
            return False

    def cleanup_stale_handles(self) -> int:
        """
        批量扫描并移除所有失效的句柄。

        Returns:
            被移除的失效句柄数量
        """
        stale_ids = []
        for workbook_id in list(self._items.keys()):
            if not self.validate_handle(workbook_id):
                stale_ids.append(workbook_id)

        for workbook_id in stale_ids:
            self._items.pop(workbook_id, None)

        if stale_ids:
            logger.info(f"[Registry] Cleaned {len(stale_ids)} stale handle(s)")
        return len(stale_ids)

    def validate_all_handles(self) -> dict:
        """
        验证所有 workbook handle，返回详细的验证结果。

        Returns:
            包含验证结果的字典，包括每个 handle 的状态
        """
        results: dict[str, dict] = {}
        valid_count = 0
        stale_count = 0

        for workbook_id, handle in list(self._items.items()):
            is_valid = self.validate_handle(workbook_id)
            results[workbook_id] = {
                "workbook_name": handle.workbook_name,
                "file_path": handle.file_path,
                "is_valid": is_valid,
                "generation": self._generation,
            }
            if is_valid:
                valid_count += 1
            else:
                stale_count += 1

        return {
            "total": len(self._items),
            "valid_count": valid_count,
            "stale_count": stale_count,
            "generation": self._generation,
            "handles": results,
        }

    def prune_stale_handles(self) -> dict:
        """
        清理所有失效的 stale handle，只做局部清理。

        Returns:
            包含清理结果的字典
        """
        before_count = len(self._items)
        stale_ids = []

        for workbook_id in list(self._items.keys()):
            if not self.validate_handle(workbook_id):
                stale_ids.append(workbook_id)

        for workbook_id in stale_ids:
            self._items.pop(workbook_id, None)

        after_count = len(self._items)
        pruned_count = before_count - after_count

        if pruned_count > 0:
            logger.info(f"[Registry] Pruned {pruned_count} stale handle(s)")

        return {
            "before_count": before_count,
            "after_count": after_count,
            "pruned_count": pruned_count,
            "pruned_ids": stale_ids,
        }

    def get_workbook_count(self) -> dict[str, int]:
        """
        返回注册统计信息。

        Returns:
            包含 registry_count, excel_count, valid_count, stale_count 的字典
        """
        excel_count = 0
        valid_count = 0
        stale_count = 0

        for handle in self._items.values():
            excel_count += 1
            if self.validate_handle(handle.workbook_id):
                valid_count += 1
            else:
                stale_count += 1

        return {
            "registry_count": len(self._items),
            "excel_count": excel_count,
            "valid_count": valid_count,
            "stale_count": stale_count,
        }

    def get_workbook(self, workbook_id: str) -> Any:
        """
        获取工作簿 COM 对象，失效时自动移除并抛出友好错误。
        """
        handle = self.require(workbook_id)
        if not self.validate_handle(workbook_id):
            self.remove(workbook_id)
            raise WorkbookHandleStaleError(workbook_id)
        return handle.workbook_obj
