"""Tests for centralized Microsoft configuration and MCP safety policy."""

from __future__ import annotations

import pytest
from fastmcp.exceptions import ToolError

from microsoft_agent.settings import (
    AuthenticationMode,
    MicrosoftSettings,
    PermissionProfile,
)
from microsoft_agent.tool_policy import (
    MicrosoftToolPolicy,
    ToolRisk,
    classify_tool_risk,
)


def test_settings_use_tenant_specific_authority_and_profiles():
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_TENANT_ID": "tenant-id",
            "MICROSOFT_CLIENT_ID": "client-id",
            "MICROSOFT_PERMISSION_PROFILES": "productivity, collaboration",
            "MICROSOFT_ENABLED_TOOL_GROUPS": "mail,calendar",
        }
    )

    assert settings.authority.endswith("/tenant-id")
    assert settings.permission_profiles == (
        PermissionProfile.PRODUCTIVITY,
        PermissionProfile.COLLABORATION,
    )
    assert "Mail.Send" in settings.graph_scopes
    assert "ChatMessage.Send" in settings.graph_scopes
    assert settings.tool_group_enabled("mail") is True
    assert settings.tool_group_enabled("devices") is False


def test_settings_application_mode_uses_default_scope(monkeypatch):
    monkeypatch.setenv("MICROSOFT_TEST_CLIENT_SECRET", "test-only-secret")
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_TENANT_ID": "tenant-id",
            "MICROSOFT_CLIENT_ID": "client-id",
            "MICROSOFT_CLIENT_SECRET_REF": "env://MICROSOFT_TEST_CLIENT_SECRET",
            "MICROSOFT_AUTH_MODE": "application",
        }
    )

    assert settings.authentication_mode is AuthenticationMode.APPLICATION
    assert settings.graph_scopes == ("https://graph.microsoft.com/.default",)
    assert "test-only-secret" not in repr(settings)


def test_raw_client_secret_variable_is_not_a_supported_configuration() -> None:
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_TENANT_ID": "tenant-id",
            "MICROSOFT_CLIENT_ID": "client-id",
            "MICROSOFT_CLIENT_SECRET": "must-not-be-consumed",
            "MICROSOFT_AUTH_MODE": "application",
        }
    )

    with pytest.raises(ValueError, match="MICROSOFT_CLIENT_SECRET_REF"):
        settings.require_identity()


def test_client_secret_reference_rejects_inline_or_unknown_schemes() -> None:
    with pytest.raises(ValueError, match="env://, vault://, or secret://"):
        MicrosoftSettings.from_env(
            {"MICROSOFT_CLIENT_SECRET_REF": "inline://must-not-be-consumed"}
        )


def test_settings_managed_identity_requires_no_stored_secret():
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_AUTH_MODE": "managed_identity",
            "MICROSOFT_MANAGED_IDENTITY_CLIENT_ID": "user-assigned-client-id",
        }
    )

    assert settings.authentication_mode is AuthenticationMode.MANAGED_IDENTITY
    assert settings.is_configured is True
    assert settings.graph_scopes == ("https://graph.microsoft.com/.default",)
    assert settings.require_identity() is settings


def test_settings_workload_identity_validates_federated_token_file(tmp_path):
    token_file = tmp_path / "federated-token"
    token_file.write_text("test-only-token", encoding="utf-8")
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_AUTH_MODE": "workload_identity",
            "MICROSOFT_TENANT_ID": "tenant-id",
            "MICROSOFT_CLIENT_ID": "client-id",
            "AZURE_FEDERATED_TOKEN_FILE": str(token_file),
        }
    )

    assert settings.authentication_mode is AuthenticationMode.WORKLOAD_IDENTITY
    assert settings.is_configured is True
    assert settings.require_identity() is settings


def test_settings_accept_exact_https_office_origins():
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_OFFICE_ADDIN_ORIGINS": (
                "https://localhost:3000,https://office.example.test"
            )
        }
    )
    assert settings.office_addin_origins == (
        "https://localhost:3000",
        "https://office.example.test",
    )


def test_settings_reject_non_https_office_origin():
    with pytest.raises(ValueError, match="HTTPS"):
        MicrosoftSettings.from_env(
            {"MICROSOFT_OFFICE_ADDIN_ORIGINS": "http://localhost:3000"}
        )


def test_settings_require_identity_reports_missing_values():
    with pytest.raises(ValueError, match="MICROSOFT_TENANT_ID"):
        MicrosoftSettings().require_identity()


def test_ingestion_pseudonymization_key_is_required_and_secret(monkeypatch):
    with pytest.raises(ValueError, match="pseudonymization key"):
        MicrosoftSettings().require_ingestion_pseudonymization_key()

    monkeypatch.setenv("MICROSOFT_TEST_PSEUDONYMIZATION_KEY", "x" * 32)
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF": (
                "env://MICROSOFT_TEST_PSEUDONYMIZATION_KEY"
            )
        }
    )
    assert settings.require_ingestion_pseudonymization_key() == b"x" * 32
    assert "x" * 32 not in repr(settings)


def test_ingestion_pseudonymization_key_rejects_short_values(monkeypatch):
    monkeypatch.setenv("MICROSOFT_TEST_PSEUDONYMIZATION_KEY", "too-short")
    settings = MicrosoftSettings.from_env(
        {
            "MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF": (
                "env://MICROSOFT_TEST_PSEUDONYMIZATION_KEY"
            )
        }
    )
    with pytest.raises(ValueError, match="at least 32 bytes"):
        settings.require_ingestion_pseudonymization_key()


@pytest.mark.parametrize(
    ("tool_name", "risk"),
    [
        ("list_mail_messages", ToolRisk.READ),
        ("send_mail", ToolRisk.WRITE),
        ("delete_mail_message", ToolRisk.DESTRUCTIVE),
        ("cancel_power_automate_desktop_flow_run", ToolRisk.DESTRUCTIVE),
        ("wipe_managed_device", ToolRisk.DESTRUCTIVE),
        ("get_word_selection_from_office", ToolRisk.READ),
        ("write_word_selection_in_office", ToolRisk.WRITE),
        ("delete_powerpoint_slide_in_office", ToolRisk.DESTRUCTIVE),
        ("list_microsoft_ingestion_projection", ToolRisk.READ),
    ],
)
def test_tool_risk_classification(tool_name, risk):
    assert classify_tool_risk(tool_name) is risk


def test_policy_is_read_only_by_default():
    policy = MicrosoftToolPolicy(MicrosoftSettings())

    assert policy.require("list_mail_messages").allowed is True
    with pytest.raises(ToolError, match="MICROSOFT_ALLOW_WRITES"):
        policy.require("send_mail")


def test_policy_separates_writes_and_destructive_actions():
    write_settings = MicrosoftSettings(allow_write_tools=True)
    write_policy = MicrosoftToolPolicy(write_settings)
    assert write_policy.require("send_mail").allowed is True
    with pytest.raises(ToolError, match="MICROSOFT_ALLOW_DESTRUCTIVE"):
        write_policy.require("delete_mail_message")

    destructive_policy = MicrosoftToolPolicy(
        MicrosoftSettings(
            allow_write_tools=True,
            allow_destructive_tools=True,
        )
    )
    assert destructive_policy.require("delete_mail_message").allowed is True
