"""Tests for audience-scoped Microsoft token acquisition."""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from microsoft_agent.auth import AuthenticationRequiredError
from microsoft_agent.integration_auth import MicrosoftAudienceTokenProvider


@pytest.mark.asyncio
async def test_audience_provider_requests_default_scope() -> None:
    manager = MagicMock()
    manager.acquire_token_for_scopes.return_value = "access-token"
    provider = MicrosoftAudienceTokenProvider(
        ["https://graph.microsoft.com/"], auth_manager=manager
    )

    token = await provider.get_token("https://graph.microsoft.com")

    assert token == "access-token"
    manager.acquire_token_for_scopes.assert_called_once_with(
        ["https://graph.microsoft.com/.default"], allow_interactive=False
    )


@pytest.mark.asyncio
async def test_audience_provider_interactive_consent_is_explicit() -> None:
    manager = MagicMock()
    manager.acquire_token_for_scopes.return_value = "access-token"
    provider = MicrosoftAudienceTokenProvider(
        ["https://contoso.crm.dynamics.com"],
        auth_manager=manager,
        allow_interactive=True,
    )

    token = await provider.get_token("https://contoso.crm.dynamics.com")

    assert token == "access-token"
    manager.acquire_token_for_scopes.assert_called_once_with(
        ["https://contoso.crm.dynamics.com/.default"], allow_interactive=True
    )


@pytest.mark.asyncio
async def test_audience_provider_denies_dynamic_resource() -> None:
    provider = MicrosoftAudienceTokenProvider(
        ["https://graph.microsoft.com"], auth_manager=MagicMock()
    )

    with pytest.raises(AuthenticationRequiredError, match="unconfigured"):
        await provider.get_token("https://example.invalid")


@pytest.mark.asyncio
async def test_audience_provider_requires_cached_authentication() -> None:
    manager = MagicMock()
    manager.acquire_token_for_scopes.return_value = None
    provider = MicrosoftAudienceTokenProvider(
        ["https://graph.microsoft.com"], auth_manager=manager
    )

    with pytest.raises(AuthenticationRequiredError, match="required"):
        await provider.get_token("https://graph.microsoft.com")


@pytest.mark.parametrize(
    "audience",
    ["http://graph.microsoft.com", "https://user:pass@example.com", "not-a-url"],
)
def test_audience_provider_rejects_unsafe_audiences(audience: str) -> None:
    with pytest.raises(ValueError, match="HTTPS"):
        MicrosoftAudienceTokenProvider([audience])


def test_audience_provider_accepts_entra_app_id_uri() -> None:
    provider = MicrosoftAudienceTokenProvider(["api://windows-companion/"])

    assert provider.allowed_audiences == frozenset({"api://windows-companion"})
