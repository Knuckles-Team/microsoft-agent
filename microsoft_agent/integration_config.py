"""Validated configuration for document, Power Platform, and laptop services."""

from __future__ import annotations

import json
from collections.abc import Mapping
from functools import lru_cache
from pathlib import Path
from typing import Any

from agent_utilities.core.config import setting
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from microsoft_agent.intune_service import IntuneServiceSettings
from microsoft_agent.power_platform import PowerPlatformSettings
from microsoft_agent.windows_companion import WindowsCompanionSettings

_MAX_CONFIG_BYTES = 1024 * 1024


class DocumentRuntimeSettings(BaseModel):
    """Filesystem trust boundaries and size limits for Office generation."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    artifact_root: Path | None = None
    template_root: Path | None = None
    max_artifact_bytes: int = Field(default=100 * 1024 * 1024, ge=1)
    max_template_bytes: int = Field(default=50 * 1024 * 1024, ge=1)


class IntegrationRuntimeSettings(BaseModel):
    """Optional non-Graph integration configuration loaded from one JSON file."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    documents: DocumentRuntimeSettings = Field(default_factory=DocumentRuntimeSettings)
    power_platform: PowerPlatformSettings | None = None
    windows_companion: WindowsCompanionSettings | None = None
    intune: IntuneServiceSettings | None = None

    @classmethod
    def from_env(
        cls, env: Mapping[str, str] | None = None
    ) -> IntegrationRuntimeSettings:
        """Load complex settings from JSON and apply simple safe env overrides."""

        raw: dict[str, Any] = {}
        config_path = _configured_value(env, "MICROSOFT_INTEGRATIONS_CONFIG_PATH")
        if config_path:
            raw = _read_json_config(Path(config_path).expanduser())

        document_raw = dict(raw.get("documents") or {})
        artifact_root = _configured_value(env, "MICROSOFT_DOCUMENT_ARTIFACT_ROOT")
        template_root = _configured_value(env, "MICROSOFT_DOCUMENT_TEMPLATE_ROOT")
        if artifact_root:
            document_raw["artifact_root"] = artifact_root
        if template_root:
            document_raw["template_root"] = template_root
        raw["documents"] = document_raw

        dataverse_url = _configured_value(env, "MICROSOFT_DATAVERSE_ENVIRONMENT_URL")
        if not raw.get("power_platform") and dataverse_url:
            named_flows = _json_env_object(
                _configured_value(
                    env, "MICROSOFT_POWER_AUTOMATE_NAMED_FLOWS_JSON"
                ),
                "MICROSOFT_POWER_AUTOMATE_NAMED_FLOWS_JSON",
            )
            raw["power_platform"] = {
                "dataverse_environment_url": dataverse_url,
                "dataverse_audience": _configured_value(
                    env, "MICROSOFT_DATAVERSE_AUDIENCE"
                ),
                "allow_lifecycle_changes": _env_bool(
                    _configured_value(
                        env,
                        "MICROSOFT_POWER_PLATFORM_ALLOW_LIFECYCLE_CHANGES",
                    )
                ),
                "named_flows": named_flows,
            }
        try:
            return cls.model_validate(raw)
        except ValidationError as exc:
            raise ValueError("Microsoft integration configuration is invalid") from exc


def _read_json_config(path: Path) -> dict[str, Any]:
    try:
        if not path.is_file():
            raise ValueError("MICROSOFT_INTEGRATIONS_CONFIG_PATH is not a file")
        if path.stat().st_size > _MAX_CONFIG_BYTES:
            raise ValueError("Microsoft integration configuration exceeds 1 MiB")
        value = json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise ValueError(
            "Microsoft integration configuration could not be read"
        ) from exc
    except json.JSONDecodeError as exc:
        raise ValueError(
            "Microsoft integration configuration is not valid JSON"
        ) from exc
    if not isinstance(value, dict):
        raise ValueError("Microsoft integration configuration must be a JSON object")
    return value


def _configured_value(
    source: Mapping[str, str] | None, name: str, default: str | None = None
) -> str | None:
    if source is None:
        return setting(name, default)
    return source.get(name, default)


def _json_env_object(raw: str | None, variable: str) -> dict[str, Any]:
    if not raw:
        return {}
    if len(raw.encode("utf-8")) > _MAX_CONFIG_BYTES:
        raise ValueError(f"{variable} exceeds 1 MiB")
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"{variable} is not valid JSON") from exc
    if not isinstance(parsed, dict):
        raise ValueError(f"{variable} must be a JSON object")
    return parsed


def _env_bool(value: str | None) -> bool:
    return bool(value and value.strip().lower() in {"1", "true", "yes", "on"})


@lru_cache(maxsize=1)
def get_integration_settings() -> IntegrationRuntimeSettings:
    """Return cached integration settings for the current process."""

    return IntegrationRuntimeSettings.from_env()


def clear_integration_settings_cache() -> None:
    """Clear the integration settings cache for tests and controlled reloads."""

    get_integration_settings.cache_clear()
