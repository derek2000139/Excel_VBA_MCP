from __future__ import annotations

from excelforge.gateway.utils import call_runtime
from excelforge.models.error_models import ErrorCode, ExcelForgeError


class _RuntimeClientOk:
    def call(self, method: str, params: dict) -> dict:
        return {"success": True, "code": "OK", "data": {"method": method, "params": params}}


class _RuntimeClientError:
    def call(self, method: str, params: dict) -> dict:
        _ = method
        _ = params
        raise ExcelForgeError(ErrorCode.E503_RUNTIME_UNAVAILABLE, "runtime down")


def test_call_runtime_passthrough_on_success() -> None:
    result = call_runtime(
        _RuntimeClientOk(),
        tool_name="workbook.open_file",
        method="workbook.open",
        params={"file_path": "D:/demo.xlsx"},
    )
    assert result["success"] is True
    assert result["data"]["method"] == "workbook.open"


def test_call_runtime_wraps_excel_forge_error() -> None:
    result = call_runtime(
        _RuntimeClientError(),
        tool_name="workbook.open_file",
        method="workbook.open",
        params={"file_path": "D:/demo.xlsx"},
    )
    assert result["success"] is False
    assert result["code"] == ErrorCode.E503_RUNTIME_UNAVAILABLE
    assert "runtime down" in result["message"]
