"""Focused regression tests for the current governed transport boundary."""

from __future__ import annotations

import base64
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

import httpx
import pytest

from microsoft_agent.api_client import MicrosoftGraphApi
from microsoft_agent.power_platform import HttpxAsyncHttpTransport
from microsoft_agent.universal_print import PrintDocumentSubmission
from microsoft_agent.windows_companion import HttpxCompanionTransport


def _tls_profile() -> SimpleNamespace:
    return SimpleNamespace(
        ssl_context=MagicMock(name="ssl-context"),
        proxy_url=None,
        cleanup=MagicMock(),
    )


def test_provider_transport_uses_pinned_tls_and_closes_once(monkeypatch) -> None:
    profile = _tls_profile()
    resolver = MagicMock(return_value=profile)
    client = MagicMock()
    factory = MagicMock(return_value=client)
    monkeypatch.setattr(
        "microsoft_agent.power_platform.resolve_configured_tls_profile", resolver
    )
    monkeypatch.setattr("microsoft_agent.power_platform.create_http_client", factory)

    transport = HttpxAsyncHttpTransport(
        service="microsoft_graph",
        tls_profile="private-ca",
        allowed_private_hosts=("configured.example",),
    )

    resolver.assert_called_once_with(
        "microsoft_graph", profile_name="private-ca", profile_ref=None
    )
    kwargs = factory.call_args.kwargs
    assert kwargs["verify"] is profile.ssl_context
    assert kwargs["pin_egress"] is True
    assert kwargs["follow_redirects"] is False
    assert kwargs["trust_env"] is False
    assert kwargs["allowed_private_hosts"] == ("configured.example",)

    transport.close()
    transport.close()
    client.close.assert_called_once_with()
    profile.cleanup.assert_called_once_with()


def test_companion_transport_uses_selected_tls_profile(monkeypatch) -> None:
    profile = _tls_profile()
    resolver = MagicMock(return_value=profile)
    client = MagicMock()
    factory = MagicMock(return_value=client)
    monkeypatch.setattr(
        "microsoft_agent.windows_companion.resolve_configured_tls_profile", resolver
    )
    monkeypatch.setattr("microsoft_agent.windows_companion.create_http_client", factory)

    transport = HttpxCompanionTransport(
        tls_profile_ref="secret://transport/companion-tls",
        allowed_private_hosts=("relay.example",),
    )

    resolver.assert_called_once_with(
        "microsoft_companion",
        profile_name=None,
        profile_ref="secret://transport/companion-tls",
    )
    assert factory.call_args.kwargs["pin_egress"] is True
    assert factory.call_args.kwargs["allowed_private_hosts"] == ("relay.example",)
    transport.close()
    client.close.assert_called_once_with()
    profile.cleanup.assert_called_once_with()


@pytest.mark.asyncio
async def test_provider_transport_sanitizes_pinned_egress_failure() -> None:
    client = MagicMock()
    client.request.side_effect = httpx.ConnectError(
        "sensitive destination must not escape"
    )
    transport = HttpxAsyncHttpTransport(service="microsoft_graph", client=client)

    with pytest.raises(OSError, match="^Provider transport failed$") as exc_info:
        await transport.request(
            "GET",
            "https://opaque.example/resource",
            headers={},
            timeout=1,
        )

    assert exc_info.value.__cause__ is None


def test_integration_cache_clear_closes_owned_transports(monkeypatch) -> None:
    from microsoft_agent import integration_tools

    integration_tools.clear_integration_client_caches()
    transport = MagicMock()
    transport_type = MagicMock(return_value=transport)
    monkeypatch.setattr(
        "microsoft_agent.integration_tools.HttpxAsyncHttpTransport", transport_type
    )

    created = integration_tools._http_transport(  # noqa: SLF001
        service="microsoft_graph",
        tls_profile="private-ca",
    )
    integration_tools.clear_integration_client_caches()

    assert created is transport
    transport_type.assert_called_once_with(
        service="microsoft_graph",
        tls_profile="private-ca",
        tls_profile_ref=None,
        allowed_private_hosts=(),
    )
    transport.close.assert_called_once_with()


def _print_api() -> MicrosoftGraphApi:
    api = object.__new__(MicrosoftGraphApi)
    api.auth_manager = SimpleNamespace(
        graph_tls_profile="private-ca",
        graph_tls_profile_ref=None,
    )
    api.create_print_job = AsyncMock(
        return_value={"id": "job-1", "documents": [{"id": "document-1"}]}
    )
    api.create_print_document_upload_session = AsyncMock(
        return_value={
            "uploadUrl": (
                "https://print.print.microsoft.com/uploadSessions/session-1"
                "?tempauthtoken=opaque"
            )
        }
    )
    api.start_print_job = AsyncMock(return_value={"state": "processing"})
    return api


def _print_submission() -> PrintDocumentSubmission:
    return PrintDocumentSubmission(
        document_name="report.pdf",
        content_type="application/pdf",
        content_base64=base64.b64encode(b"pdf").decode("ascii"),
    )


@pytest.mark.asyncio
async def test_print_upload_closes_only_its_default_transport() -> None:
    api = _print_api()
    owned_transport = MagicMock()
    transport_type = MagicMock(return_value=owned_transport)
    uploader = MagicMock()
    uploader.upload = AsyncMock(return_value={"id": "document-1"})

    with (
        patch(
            "microsoft_agent.power_platform.HttpxAsyncHttpTransport",
            transport_type,
        ),
        patch(
            "microsoft_agent.universal_print.UniversalPrintUploader",
            MagicMock(return_value=uploader),
        ),
    ):
        result = await api.submit_print_document("printer-1", _print_submission())

    assert result["document"]["id"] == "document-1"
    transport_type.assert_called_once_with(
        service="microsoft_graph",
        tls_profile="private-ca",
        tls_profile_ref=None,
    )
    owned_transport.close.assert_called_once_with()

    injected_transport = MagicMock()
    uploader.upload.reset_mock(return_value=True, side_effect=True)
    uploader.upload.return_value = {"id": "document-1"}
    with (
        patch(
            "microsoft_agent.power_platform.HttpxAsyncHttpTransport",
            transport_type,
        ),
        patch(
            "microsoft_agent.universal_print.UniversalPrintUploader",
            MagicMock(return_value=uploader),
        ),
    ):
        await api.submit_print_document(
            "printer-1",
            _print_submission(),
            upload_transport=injected_transport,
        )

    assert transport_type.call_count == 1
    injected_transport.close.assert_not_called()
