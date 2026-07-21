"""Tests for shared authentication and HTTP protocol adapters."""

from __future__ import annotations

import json
from collections.abc import Mapping
from typing import Any

import pytest

from microsoft_agent.integration_adapters import (
    IntuneGraphTokenAdapter,
    IntuneHttpClientAdapter,
)
from microsoft_agent.intune_service import GRAPH_AUDIENCE
from microsoft_agent.power_platform import HttpResponse


class AudienceProvider:
    def __init__(self) -> None:
        self.requested: list[str] = []

    async def get_token(self, audience: str) -> str:
        self.requested.append(audience)
        return "token-value"


class Transport:
    def __init__(self) -> None:
        self.request_data: dict[str, Any] | None = None

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        timeout: float,
    ) -> HttpResponse:
        self.request_data = {
            "method": method,
            "url": url,
            "headers": headers,
            "params": params,
            "body": body,
            "timeout": timeout,
        }
        return HttpResponse(
            status_code=200,
            headers={"Content-Type": "application/json"},
            body=b'{"value": []}',
        )


@pytest.mark.asyncio
async def test_intune_token_adapter_attests_graph_audience() -> None:
    provider = AudienceProvider()
    adapter = IntuneGraphTokenAdapter(provider)  # type: ignore[arg-type]

    token = await adapter.get_token(GRAPH_AUDIENCE)

    assert token.access_token.get_secret_value() == "token-value"
    assert token.audience == GRAPH_AUDIENCE
    assert provider.requested == [GRAPH_AUDIENCE]


@pytest.mark.asyncio
async def test_intune_http_adapter_serializes_json() -> None:
    transport = Transport()
    adapter = IntuneHttpClientAdapter(  # type: ignore[arg-type]
        transport, timeout_seconds=12
    )

    response = await adapter.request(
        "POST",
        "https://graph.microsoft.com/v1.0/example",
        headers={"Authorization": "Bearer test"},
        json={"quickScan": True},
    )

    assert response.json() == {"value": []}
    assert transport.request_data is not None
    assert json.loads(transport.request_data["body"]) == {"quickScan": True}
    assert transport.request_data["timeout"] == 12
