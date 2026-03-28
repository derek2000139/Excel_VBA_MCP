from __future__ import annotations

import json
import threading
from typing import Callable


class JsonRpcPipeServer:
    def __init__(
        self,
        *,
        pipe_name: str,
        request_handler: Callable[[dict], dict],
        stop_event: threading.Event,
    ) -> None:
        self._pipe_name = pipe_name
        self._request_handler = request_handler
        self._stop_event = stop_event
        self._threads: list[threading.Thread] = []

    def serve_forever(self) -> None:
        try:
            import pywintypes  # type: ignore
            import win32file  # type: ignore
            import win32pipe  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise RuntimeError(f"pywin32 is required for runtime pipe server: {exc}") from exc

        pipe_mode = win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT

        while not self._stop_event.is_set():
            pipe = win32pipe.CreateNamedPipe(
                self._pipe_name,
                win32pipe.PIPE_ACCESS_DUPLEX,
                pipe_mode,
                win32pipe.PIPE_UNLIMITED_INSTANCES,
                1024 * 1024,
                1024 * 1024,
                1000,
                None,
            )
            connected = False
            try:
                try:
                    win32pipe.ConnectNamedPipe(pipe, None)
                except pywintypes.error as exc:
                    if exc.winerror != 535:
                        raise
                connected = True
                thread = threading.Thread(
                    target=self._handle_client,
                    args=(pipe,),
                    name="runtime-pipe-client",
                    daemon=True,
                )
                thread.start()
                self._threads.append(thread)
            finally:
                if not connected:
                    win32file.CloseHandle(pipe)

    def _handle_client(self, pipe: int) -> None:
        try:
            import pywintypes  # type: ignore
            import win32file  # type: ignore
            import win32pipe  # type: ignore
        except Exception:
            return

        buffer = b""
        try:
            while not self._stop_event.is_set():
                try:
                    _, chunk = win32file.ReadFile(pipe, 4096)
                except pywintypes.error as exc:
                    if exc.winerror in {109, 233}:
                        break
                    raise

                if not chunk:
                    break
                buffer += bytes(chunk)
                while b"\n" in buffer:
                    line, buffer = buffer.split(b"\n", 1)
                    line = line.strip()
                    if not line:
                        continue
                    response = self._safe_handle_payload(line)
                    payload = (json.dumps(response, ensure_ascii=False) + "\n").encode("utf-8")
                    win32file.WriteFile(pipe, payload)
        finally:
            try:
                win32pipe.DisconnectNamedPipe(pipe)
            except Exception:
                pass
            try:
                win32file.CloseHandle(pipe)
            except Exception:
                pass

    def _safe_handle_payload(self, payload: bytes) -> dict:
        try:
            request = json.loads(payload.decode("utf-8"))
            if not isinstance(request, dict):
                raise ValueError("request is not an object")
            return self._request_handler(request)
        except Exception as exc:
            return {
                "jsonrpc": "2.0",
                "id": None,
                "error": {
                    "code": -32700,
                    "message": "invalid_request",
                    "data": {"detail": str(exc)},
                },
            }
