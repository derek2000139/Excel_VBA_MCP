from __future__ import annotations

from dataclasses import dataclass

from excelforge.gateway.runtime_client import RuntimeClient


@dataclass
class GatewayToolContext:
    runtime: RuntimeClient
