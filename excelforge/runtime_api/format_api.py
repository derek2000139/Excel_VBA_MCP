from __future__ import annotations

from typing import Any

from excelforge.runtime_api.context import RuntimeApiContext


class FormatApi:
    def __init__(self, ctx: RuntimeApiContext) -> None:
        self._ctx = ctx

    def set_style(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        range_address = str(params.get("range", ""))

        if params.get("style"):
            style = params.get("style")
        else:
            style = {}
            if "number_format" in params and params["number_format"]:
                style["number_format"] = params["number_format"]
            if "name" in params:
                style["font_name"] = params["name"]
            if "size" in params:
                style["font_size"] = params["size"]
            if "bold" in params:
                style["font_bold"] = params["bold"]
            if "italic" in params:
                style["font_italic"] = params["italic"]
            if "color" in params:
                pass
            if "font_color" in params:
                style["font_color"] = params["font_color"]
            if "fill_color" in params:
                style["fill_color"] = params["fill_color"]
            if "pattern" in params:
                pass
            if "border_style" in params:
                style["border_style"] = params["border_style"]
            if "border_type" in params:
                style["border_type"] = params["border_type"]
            if "horizontal" in params:
                style["horizontal_alignment"] = params["horizontal"]
            if "vertical" in params:
                style["vertical_alignment"] = params["vertical"]
            if "wrap_text" in params:
                style["wrap_text"] = params["wrap_text"]

        return self._ctx.run_operation(
            method_name="format.set_style",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.format_service.set_range_style(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                range_address=range_address,
                style=style,
            ),
            args_summary={
                "workbook_id": workbook_id,
                "sheet_name": sheet_name,
                "range": range_address,
            },
            default_workbook_id=workbook_id,
        )

    def auto_fit(self, params: dict[str, Any], actor_id: str) -> dict[str, Any]:
        workbook_id = str(params.get("workbook_id", ""))
        sheet_name = str(params.get("sheet_name", ""))
        columns = params.get("range")
        return self._ctx.run_operation(
            method_name="format.auto_fit",
            actor_id=actor_id,
            client_request_id=params.get("client_request_id"),
            operation_fn=lambda: self._ctx.services.format_service.auto_fit_columns(
                workbook_id=workbook_id,
                sheet_name=sheet_name,
                columns=columns,
            ),
            args_summary={"workbook_id": workbook_id, "sheet_name": sheet_name, "range": columns},
            default_workbook_id=workbook_id,
        )
