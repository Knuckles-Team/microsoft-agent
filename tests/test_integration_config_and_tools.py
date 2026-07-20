"""Tests for integration configuration and modular MCP registration."""

from __future__ import annotations

import json

import pytest
from fastmcp import FastMCP

from microsoft_agent.integration_config import IntegrationRuntimeSettings
from microsoft_agent.integration_tools import (
    register_document_tools,
    register_intune_tools,
    register_power_platform_tools,
    register_windows_companion_tools,
)


def test_integration_config_loads_complex_json_and_env_overrides(tmp_path) -> None:
    config = tmp_path / "integrations.json"
    config.write_text(
        json.dumps(
            {
                "power_platform": {
                    "dataverse_environment_url": "https://contoso.crm.dynamics.com",
                    "named_flows": {
                        "daily-report": {
                            "trigger_url": "https://prod-00.westus.logic.azure.com/trigger",
                            "audience": "https://service.flow.microsoft.com/",
                        }
                    },
                }
            }
        ),
        encoding="utf-8",
    )

    settings = IntegrationRuntimeSettings.from_env(
        {
            "MICROSOFT_INTEGRATIONS_CONFIG_PATH": str(config),
            "MICROSOFT_DOCUMENT_ARTIFACT_ROOT": str(tmp_path / "artifacts"),
        }
    )

    assert settings.documents.artifact_root == tmp_path / "artifacts"
    assert settings.power_platform is not None
    assert "daily-report" in settings.power_platform.named_flows


def test_integration_config_rejects_unknown_fields(tmp_path) -> None:
    config = tmp_path / "integrations.json"
    config.write_text('{"unexpected": true}', encoding="utf-8")

    with pytest.raises(ValueError, match="invalid"):
        IntegrationRuntimeSettings.from_env(
            {"MICROSOFT_INTEGRATIONS_CONFIG_PATH": str(config)}
        )


def test_integration_config_supports_dataverse_env() -> None:
    settings = IntegrationRuntimeSettings.from_env(
        {
            "MICROSOFT_DATAVERSE_ENVIRONMENT_URL": ("https://contoso.crm.dynamics.com"),
            "MICROSOFT_POWER_AUTOMATE_NAMED_FLOWS_JSON": "{}",
        }
    )
    assert settings.power_platform is not None
    assert settings.power_platform.token_audience == (
        "https://contoso.crm.dynamics.com"
    )


@pytest.mark.asyncio
async def test_all_integration_tool_families_register() -> None:
    mcp = FastMCP("integration-test")
    register_document_tools(mcp)
    register_power_platform_tools(mcp)
    register_windows_companion_tools(mcp)
    register_intune_tools(mcp)

    names = {tool.name for tool in await mcp.list_tools()}

    assert {
        "generate_word_document",
        "generate_powerpoint_presentation",
        "generate_and_upload_word_document",
        "list_power_automate_flows",
        "run_power_automate_flow",
        "get_windows_device_health",
        "submit_windows_action",
        "list_intune_managed_devices",
        "sync_intune_device",
    } <= names
