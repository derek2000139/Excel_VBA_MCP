from __future__ import annotations

import time
from typing import Any

from excelforge.models.error_models import ErrorCode, ExcelForgeError


def gateway_error_envelope(tool_name: str, code: ErrorCode | str, message: str, duration_ms: int) -> dict[str, Any]:
    code_str = code.value if isinstance(code, ErrorCode) else str(code)
    return {
        "success": False,
        "code": code_str,
        "message": message,
        "data": None,
        "meta": {
            "tool_name": tool_name,
            "operation_id": "op_gateway",
            "workbook_id": None,
            "snapshot_id": None,
            "rollback_supported": False,
            "duration_ms": duration_ms,
            "server_version": "2.0.0",
            "client_request_id": None,
            "warnings": [],
        },
    }


def call_runtime(runtime_client, *, tool_name: str, method: str, params: dict[str, Any]) -> dict[str, Any]:
    import logging
    logger = logging.getLogger(__name__)
    logger.info("[Gateway] Calling tool=%s method=%s params=%s", tool_name, method, params)

    started = time.perf_counter()
    try:
        result = runtime_client.call(method, params)
        duration_ms = int((time.perf_counter() - started) * 1000)
        logger.info("[Gateway] tool=%s completed in %dms success=%s", tool_name, duration_ms, result.get("success") if isinstance(result, dict) else "unknown")
        return result
    except ExcelForgeError as exc:
        duration_ms = int((time.perf_counter() - started) * 1000)
        logger.error("[Gateway] tool=%s failed: %s", tool_name, exc.message)
        return gateway_error_envelope(tool_name=tool_name, code=exc.code, message=exc.message, duration_ms=duration_ms)
    except Exception as exc:  # noqa: BLE001
        duration_ms = int((time.perf_counter() - started) * 1000)
        logger.error("[Gateway] tool=%s exception: %s", tool_name, exc)
        return gateway_error_envelope(
            tool_name=tool_name,
            code=ErrorCode.E500_INTERNAL,
            message=str(exc),
            duration_ms=duration_ms,
        )
