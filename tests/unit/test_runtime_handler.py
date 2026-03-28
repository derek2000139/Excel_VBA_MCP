from __future__ import annotations

from excelforge.models.error_models import ErrorCode, ExcelForgeError
from excelforge.runtime.handler import RuntimeJsonRpcHandler


class _OkDispatcher:
    def dispatch(self, method: str, params: dict, actor_id: str) -> dict:
        return {
            "success": True,
            "code": "OK",
            "message": "operation completed",
            "data": {"method": method, "actor_id": actor_id, "params": params},
            "meta": {
                "tool_name": method,
                "operation_id": "op_test",
                "workbook_id": None,
                "snapshot_id": None,
                "rollback_supported": False,
                "duration_ms": 1,
                "server_version": "2.0.0",
                "client_request_id": params.get("client_request_id"),
                "warnings": [],
            },
        }


class _ErrorDispatcher:
    def dispatch(self, method: str, params: dict, actor_id: str) -> dict:
        _ = method
        _ = params
        _ = actor_id
        raise ExcelForgeError(ErrorCode.E404_WORKBOOK_NOT_OPEN, "Workbook not open")


def test_runtime_handler_success_response_shape() -> None:
    handler = RuntimeJsonRpcHandler(_OkDispatcher())  # type: ignore[arg-type]
    req = {
        "jsonrpc": "2.0",
        "id": "req_1",
        "method": "workbook.open",
        "params": {
            "file_path": "D:/demo.xlsx",
            "actor_id": "excel-core-mcp",
            "client_request_id": "cid-1",
        },
    }
    resp = handler.handle_request(req)
    assert resp["jsonrpc"] == "2.0"
    assert resp["id"] == "req_1"
    assert "result" in resp
    assert resp["result"]["success"] is True
    assert resp["result"]["data"]["actor_id"] == "excel-core-mcp"


def test_runtime_handler_error_response_shape() -> None:
    handler = RuntimeJsonRpcHandler(_ErrorDispatcher())  # type: ignore[arg-type]
    req = {
        "jsonrpc": "2.0",
        "id": "req_2",
        "method": "workbook.info",
        "params": {"workbook_id": "wb_1", "actor_id": "excel-core-mcp"},
    }
    resp = handler.handle_request(req)
    assert resp["jsonrpc"] == "2.0"
    assert resp["id"] == "req_2"
    assert "error" in resp
    assert resp["error"]["data"]["error_code"] == ErrorCode.E404_WORKBOOK_NOT_OPEN
