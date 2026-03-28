from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml


class ProfileResolutionError(Exception):
    pass


class ProfileResolver:
    def __init__(self, profiles_path: str | Path | None = None) -> None:
        if profiles_path is None:
            profiles_path = Path(__file__).parent / "profiles.yaml"
        self._profiles_path = Path(profiles_path)
        self._profiles: dict[str, Any] = {}
        self._load_profiles()

    def _load_profiles(self) -> None:
        if not self._profiles_path.exists():
            raise ProfileResolutionError(f"Profiles file not found: {self._profiles_path}")
        data = yaml.safe_load(self._profiles_path.read_text(encoding="utf-8")) or {}
        self._profiles = data.get("profiles", {})
        if not self._profiles:
            raise ProfileResolutionError(f"No profiles found in {self._profiles_path}")

    def resolve(self, profile_name: str) -> dict[str, Any]:
        if profile_name not in self._profiles:
            available = ", ".join(sorted(self._profiles.keys()))
            raise ProfileResolutionError(
                f"Profile '{profile_name}' not found. Available profiles: {available}"
            )
        profile = self._profiles[profile_name]
        return {
            "name": profile_name,
            "description": profile.get("description", ""),
            "bundles": list(profile.get("bundles", [])),
            "tool_budget": profile.get("tool_budget"),
            "risk_level": profile.get("risk_level", "low"),
            "is_development": profile.get("is_development", False),
        }

    def list_profiles(self) -> list[str]:
        return sorted(self._profiles.keys())

    def get_profile_info(self, profile_name: str) -> dict[str, Any]:
        if profile_name not in self._profiles:
            raise ProfileResolutionError(f"Profile '{profile_name}' not found")
        profile = self._profiles[profile_name]
        return {
            "name": profile_name,
            "description": profile.get("description", ""),
            "bundles": list(profile.get("bundles", [])),
            "tool_budget": profile.get("tool_budget"),
            "risk_level": profile.get("risk_level", "low"),
            "is_development": profile.get("is_development", False),
        }

    def validate_bundles(self, bundle_names: list[str], bundle_registry: "BundleRegistry") -> None:
        available_bundles = bundle_registry.list_bundles()
        for bundle_name in bundle_names:
            if bundle_name not in available_bundles:
                raise ProfileResolutionError(
                    f"Bundle '{bundle_name}' in profile not found. Available bundles: {available_bundles}"
                )


class BundleRegistry:
    def __init__(self, bundles_path: str | Path | None = None) -> None:
        if bundles_path is None:
            bundles_path = Path(__file__).parent / "bundles.yaml"
        self._bundles_path = Path(bundles_path)
        self._bundles: dict[str, Any] = {}
        self._domains: dict[str, Any] = {}
        self._load_bundles()

    def _load_bundles(self) -> None:
        if not self._bundles_path.exists():
            raise ProfileResolutionError(f"Bundles file not found: {self._bundles_path}")
        data = yaml.safe_load(self._bundles_path.read_text(encoding="utf-8")) or {}
        self._bundles = data.get("bundles", {})
        self._domains = data.get("domains", {})
        if not self._bundles:
            raise ProfileResolutionError(f"No bundles found in {self._bundles_path}")

    def resolve_bundles(self, bundle_names: list[str]) -> list[str]:
        resolved: list[str] = []
        visited: set[str] = set()

        def visit(name: str) -> None:
            if name in visited:
                return
            visited.add(name)
            if name not in self._bundles:
                raise ProfileResolutionError(f"Bundle '{name}' not found")
            bundle = self._bundles[name]
            for dep in bundle.get("dependencies", []):
                visit(dep)
            if name not in resolved:
                resolved.append(name)

        for name in bundle_names:
            visit(name)
        return resolved

    def get_bundle_tools(self, bundle_name: str) -> list[str]:
        if bundle_name not in self._bundles:
            raise ProfileResolutionError(f"Bundle '{bundle_name}' not found")
        bundle = self._bundles[bundle_name]
        domain_names = bundle.get("domains", [])
        tools: list[str] = []
        for domain_name in domain_names:
            if domain_name in self._domains:
                tools.extend(self._domains[domain_name].get("tools", []))
        return tools

    def get_all_tools(self, bundle_names: list[str]) -> list[str]:
        resolved_bundles = self.resolve_bundles(bundle_names)
        tools: list[str] = []
        seen: set[str] = set()
        for bundle_name in resolved_bundles:
            bundle_tools = self.get_bundle_tools(bundle_name)
            for tool in bundle_tools:
                if tool not in seen:
                    seen.add(tool)
                    tools.append(tool)
        return tools

    def list_bundles(self) -> list[str]:
        return sorted(self._bundles.keys())

    def get_bundle_info(self, bundle_name: str) -> dict[str, Any]:
        if bundle_name not in self._bundles:
            raise ProfileResolutionError(f"Bundle '{bundle_name}' not found")
        bundle = self._bundles[bundle_name]
        return {
            "name": bundle_name,
            "description": bundle.get("description", ""),
            "domains": list(bundle.get("domains", [])),
            "dependencies": list(bundle.get("dependencies", [])),
            "recommended_with": bundle.get("recommended_with", []),
        }

    def list_domains(self) -> list[str]:
        return sorted(self._domains.keys())

    def get_domain_tools(self, domain_name: str) -> list[str]:
        if domain_name not in self._domains:
            raise ProfileResolutionError(f"Domain '{domain_name}' not found")
        return list(self._domains[domain_name].get("tools", []))
