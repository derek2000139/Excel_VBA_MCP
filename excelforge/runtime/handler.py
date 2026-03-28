from __future__ import annotations

from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError, normalize_exception
from excelforge.runtime_api import RuntimeApiDispatcher


class RuntimeJsonRpcHandler:
    def __init__(self, dispatcher: RuntimeApiDispatcher) -> None:
        self._dispatcher = dispatcher

    def handle_request(self, request: dict[str, Any]) -> dict[str, Any]:
        request_id = request.get("id")
        try:
            if request.get("jsonrpc") != "2.0":
                raise ExcelForgeError(ErrorCode.E400_BAD_REQUEST, "jsonrpc must be '2.0'")
            method = str(request.get("method", ""))
            if not method:
                raise ExcelForgeError(ErrorCode.E400_BAD_REQUEST, "method is required")
            params_raw = request.get("params") or {}
            if not isinstance(params_raw, dict):
                raise ExcelForgeError(ErrorCode.E400_BAD_REQUEST, "params must be an object")

            actor_id = str(params_raw.get("actor_id", "unknown-gateway"))
            result = self._dispatcher.dispatch(method=method, params=params_raw, actor_id=actor_id)
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": result,
            }
        except Exception as exc:  # noqa: BLE001
            err = normalize_exception(exc)
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32000,
                    "message": err.code.value,
                    "data": {
                        "error_code": err.code.value,
                        "detail": err.message,
                        "suggestion": self._suggestion_for(err.code),
                    },
                },
            }

    @staticmethod
    def _suggestion_for(code: ErrorCode) -> str:
        if code == ErrorCode.E503_RUNTIME_UNAVAILABLE:
            return "Ensure runtime process is started and runtime.lock points to a live process"
        if code == ErrorCode.E503_RUNTIME_TIMEOUT:
            return "Retry the operation or increase call timeout in gateway config"
        if code == ErrorCode.E404_WORKBOOK_NOT_OPEN:
            return "Open workbook first with workbook.open"
        return "Review request parameters and runtime logs"
