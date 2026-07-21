"""Tests for offline sanitized configuration validation."""

from __future__ import annotations

import json

from microsoft_agent.auth import clear_auth_manager_cache
from microsoft_agent.config_cli import configuration_report, main
from microsoft_agent.integration_config import clear_integration_settings_cache
from microsoft_agent.settings import clear_settings_cache


def clear_caches() -> None:
    clear_settings_cache()
    clear_integration_settings_cache()
    clear_auth_manager_cache()


def test_configuration_report_works_before_authentication(monkeypatch) -> None:
    monkeypatch.delenv("MICROSOFT_TENANT_ID", raising=False)
    monkeypatch.delenv("MICROSOFT_CLIENT_ID", raising=False)
    clear_caches()
    try:
        report = configuration_report()
    finally:
        clear_caches()

    assert report["identity"]["configured"] is False
    assert isinstance(report["documents"]["word_available"], bool)


def test_configuration_cli_never_prints_secret(monkeypatch, capsys) -> None:
    monkeypatch.setenv("MICROSOFT_TENANT_ID", "tenant-id")
    monkeypatch.setenv("MICROSOFT_CLIENT_ID", "client-id")
    monkeypatch.setenv("MICROSOFT_AUTH_MODE", "application")
    monkeypatch.setenv("MICROSOFT_TEST_CLIENT_SECRET", "never-print-this-secret")
    monkeypatch.setenv(
        "MICROSOFT_CLIENT_SECRET_REF", "env://MICROSOFT_TEST_CLIENT_SECRET"
    )
    clear_caches()
    try:
        code = main(["--require-identity"])
    finally:
        clear_caches()

    output = capsys.readouterr().out
    assert code == 0
    assert "never-print-this-secret" not in output
    assert json.loads(output)["identity"]["configured"] is True
