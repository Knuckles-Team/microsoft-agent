"""Microsoft Graph shared TLS transport contract."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class _ConcreteApi(MicrosoftGraphApiBase):
    def verify_login(self) -> str:
        return "Authenticated with workload identity"


def _manager() -> SimpleNamespace:
    return SimpleNamespace(
        scopes=["https://graph.microsoft.com/.default"],
        graph_base_url="https://graph.microsoft.com/v1.0",
        graph_tls_profile="private-ca",
        graph_tls_profile_ref=None,
    )


def test_graph_transport_uses_shared_tls_profile_without_redirects(monkeypatch) -> None:
    profile = MagicMock()
    profile.httpx_kwargs.return_value = {
        "verify": MagicMock(name="ssl-context"),
        "trust_env": False,
    }
    resolver = MagicMock(return_value=profile)
    client = MagicMock()
    client.aclose = AsyncMock()
    httpx_client = MagicMock(return_value=client)
    adapter = MagicMock()
    adapter_type = MagicMock(return_value=adapter)
    graph_client = MagicMock()
    graph_client_type = MagicMock(return_value=graph_client)
    auth_provider_type = MagicMock(return_value=MagicMock())

    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.resolve_configured_tls_profile", resolver
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.create_async_http_client", httpx_client
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.AzureIdentityAuthenticationProvider",
        auth_provider_type,
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.GraphRequestAdapter", adapter_type
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.GraphServiceClient", graph_client_type
    )

    api = _ConcreteApi(_manager())  # type: ignore[arg-type]

    resolver.assert_called_once_with(
        "microsoft_graph", profile_name="private-ca", profile_ref=None
    )
    kwargs = httpx_client.call_args.kwargs
    assert kwargs["follow_redirects"] is False
    assert kwargs["trust_env"] is False
    assert kwargs["verify"] is profile.httpx_kwargs.return_value["verify"]
    assert kwargs["pin_egress"] is True
    assert kwargs["allowed_private_hosts"] == ()
    adapter_type.assert_called_once_with(
        auth_provider_type.return_value,
        client=client,
    )
    assert adapter.base_url == "https://graph.microsoft.com/v1.0"
    graph_client_type.assert_called_once_with(request_adapter=adapter)
    assert api.client is graph_client


@pytest.mark.asyncio
async def test_graph_transport_closes_client_and_tls_material(monkeypatch) -> None:
    profile = MagicMock()
    profile.httpx_kwargs.return_value = {"verify": True, "trust_env": False}
    client = MagicMock()
    client.aclose = AsyncMock()
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.resolve_configured_tls_profile",
        MagicMock(return_value=profile),
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.create_async_http_client",
        MagicMock(return_value=client),
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.AzureIdentityAuthenticationProvider",
        MagicMock(),
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.GraphRequestAdapter", MagicMock()
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.GraphServiceClient", MagicMock()
    )

    api = _ConcreteApi(_manager())  # type: ignore[arg-type]
    await api.close()

    client.aclose.assert_awaited_once_with()
    profile.cleanup.assert_called_once_with()

    await api.close()
    client.aclose.assert_awaited_once_with()
    profile.cleanup.assert_called_once_with()


def test_graph_transport_cleans_tls_when_client_creation_fails(monkeypatch) -> None:
    profile = MagicMock()
    profile.httpx_kwargs.return_value = {"verify": True, "trust_env": False}
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.resolve_configured_tls_profile",
        MagicMock(return_value=profile),
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.create_async_http_client",
        MagicMock(side_effect=RuntimeError("creation failed")),
    )
    monkeypatch.setattr(
        "microsoft_agent.api.api_client_base.AzureIdentityAuthenticationProvider",
        MagicMock(),
    )

    with pytest.raises(RuntimeError, match="creation failed"):
        _ConcreteApi(_manager())  # type: ignore[arg-type]

    profile.cleanup.assert_called_once_with()


@pytest.mark.parametrize(
    "endpoint",
    [
        "http://graph.microsoft.com/v1.0",
        "https://user:password@graph.microsoft.com/v1.0",
        "https://graph.microsoft.com/v1.0?tenant=value",
    ],
)
def test_graph_endpoint_rejects_insecure_authority(endpoint: str) -> None:
    from microsoft_agent.settings import MicrosoftSettings

    with pytest.raises(ValueError, match="HTTPS URL"):
        MicrosoftSettings(graph_base_url=endpoint)
