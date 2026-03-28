from __future__ import annotations

import argparse
import signal
import threading

from excelforge.runtime.bootstrap import create_runtime_services
from excelforge.runtime.handler import RuntimeJsonRpcHandler
from excelforge.runtime.lifecycle import remove_runtime_lock, write_runtime_lock
from excelforge.runtime.pipe_server import JsonRpcPipeServer
from excelforge.runtime_api import RuntimeApiContext, RuntimeApiDispatcher


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="excelforge.runtime", description="ExcelForge Runtime Service")
    parser.add_argument("--config", default=None, help="Path to runtime-config.yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    services = create_runtime_services(args.config)
    ctx = RuntimeApiContext(services)
    dispatcher = RuntimeApiDispatcher(ctx)
    services.server_service.set_tool_names(dispatcher.method_names())
    handler = RuntimeJsonRpcHandler(dispatcher)

    stop_event = threading.Event()

    def _shutdown(*_: object) -> None:
        stop_event.set()

    signal.signal(signal.SIGINT, _shutdown)
    signal.signal(signal.SIGTERM, _shutdown)

    write_runtime_lock(services.config, args.config)
    server = JsonRpcPipeServer(
        pipe_name=services.config.runtime.pipe_name,
        request_handler=handler.handle_request,
        stop_event=stop_event,
    )

    try:
        server.serve_forever()
    finally:
        remove_runtime_lock(services.config)
        services.shutdown()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
