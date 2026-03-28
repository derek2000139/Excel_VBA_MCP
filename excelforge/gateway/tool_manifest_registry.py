from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class ToolManifest:
    name: str
    domain: str
    bundle: str
    runtime_method: str
    requires_excel_ready: bool = True
    legacy_entrypoints: list[str] = field(default_factory=list)
    maturity: str = "stable"
    risk_level: str = "low"
    description: str = ""


class ToolManifestRegistry:
    def __init__(self, bundles_path: str | Path | None = None) -> None:
        if bundles_path is None:
            bundles_path = Path(__file__).parent / "bundles.yaml"
        self._bundles_path = Path(bundles_path)
        self._manifests: dict[str, ToolManifest] = {}
        self._load_manifests()

    def _load_manifests(self) -> None:
        if not self._bundles_path.exists():
            return
        data = yaml.safe_load(self._bundles_path.read_text(encoding="utf-8")) or {}
        domains = data.get("domains", {})
        bundles = data.get("bundles", {})

        bundle_of_domain: dict[str, str] = {}
        for bundle_name, bundle_data in bundles.items():
            for domain_name in bundle_data.get("domains", []):
                bundle_of_domain[domain_name] = bundle_name

        for domain_name, domain_data in domains.items():
            bundle_name = bundle_of_domain.get(domain_name, "unknown")
            for tool_name in domain_data.get("tools", []):
                self._manifests[tool_name] = ToolManifest(
                    name=tool_name,
                    domain=domain_name,
                    bundle=bundle_name,
                    runtime_method=tool_name,
                    requires_excel_ready=True,
                    maturity=domain_data.get("maturity", "stable"),
                    risk_level=domain_data.get("risk_level", "low"),
                    legacy_entrypoints=self._infer_legacy_entrypoints(tool_name),
                )

    def _infer_legacy_entrypoints(self, tool_name: str) -> list[str]:
        entrypoint_map = {
            "server": ["core"],
            "workbook": ["core"],
            "names": ["core"],
            "sheet": ["core"],
            "range": ["core"],
            "formula": ["core"],
            "format": ["core"],
            "vba": ["vba"],
            "recovery": ["recovery"],
            "pq": ["pq"],
            "analysis": ["core"],
        }
        parts = tool_name.split(".")
        if parts:
            prefix = parts[0]
            return entrypoint_map.get(prefix, ["core"])
        return ["core"]

    def get_manifest(self, tool_name: str) -> ToolManifest | None:
        return self._manifests.get(tool_name)

    def list_tools(self) -> list[str]:
        return sorted(self._manifests.keys())

    def list_tools_by_bundle(self, bundle_name: str) -> list[str]:
        return [
            name
            for name, manifest in self._manifests.items()
            if manifest.bundle == bundle_name
        ]

    def list_tools_by_domain(self, domain_name: str) -> list[str]:
        return [
            name
            for name, manifest in self._manifests.items()
            if manifest.domain == domain_name
        ]

    def filter_tools(
        self,
        tools: list[str],
        maturity: str | None = None,
        risk_level: str | None = None,
    ) -> list[str]:
        result = []
        for tool_name in tools:
            manifest = self._manifests.get(tool_name)
            if manifest is None:
                continue
            if maturity is not None and manifest.maturity != maturity:
                continue
            if risk_level is not None and manifest.risk_level != risk_level:
                continue
            result.append(tool_name)
        return result

    def to_dict(self, tool_name: str) -> dict[str, Any] | None:
        manifest = self.get_manifest(tool_name)
        if manifest is None:
            return None
        return {
            "name": manifest.name,
            "domain": manifest.domain,
            "bundle": manifest.bundle,
            "runtime_method": manifest.runtime_method,
            "requires_excel_ready": manifest.requires_excel_ready,
            "legacy_entrypoints": manifest.legacy_entrypoints,
            "maturity": manifest.maturity,
            "risk_level": manifest.risk_level,
            "description": manifest.description,
        }
