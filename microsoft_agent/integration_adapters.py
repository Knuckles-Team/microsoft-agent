"""Small adapters between shared authentication/transport and service protocols."""

from __future__ import annotations

import json as jsonlib
from collections.abc import Mapping
from typing import Any

from microsoft_agent.integration_auth import MicrosoftAudienceTokenProvider
from microsoft_agent.intune_service import (
    GRAPH_AUDIENCE,
    GraphAccessToken,
    HttpResponse,
)
from microsoft_agent.power_platform import AsyncHttpTransport


class IntuneGraphTokenAdapter:
    """Adapt the shared audience provider to Intune's attested token model."""

    def __init__(self, provider: MicrosoftAudienceTokenProvider) -> None:
        self._provider = provider

    async def get_token(self, audience: str) -> GraphAccessToken:
        """Acquire and wrap a token whose target is fixed to Microsoft Graph."""

        if audience.rstrip("/") != GRAPH_AUDIENCE:
            raise ValueError("Intune token audience must be Microsoft Graph")
        token = await self._provider.get_token(GRAPH_AUDIENCE)
        return GraphAccessToken(access_token=token, audience=GRAPH_AUDIENCE)


class IntuneHttpResponseAdapter:
    """Expose a shared transport response through Intune's response protocol."""

    def __init__(self, status_code: int, headers: Mapping[str, str], body: bytes):
        self.status_code = status_code
        self.headers: Mapping[str, str] = dict(headers)
        self._body = body

    def json(self) -> Any:
        """Decode the buffered response body as JSON."""

        if not self._body:
            return {}
        return jsonlib.loads(self._body.decode("utf-8"))


class IntuneHttpClientAdapter:
    """Adapt the common no-redirect async HTTP transport for Intune calls."""

    def __init__(self, transport: AsyncHttpTransport, *, timeout_seconds: float = 60):
        self._transport = transport
        self._timeout_seconds = timeout_seconds

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, str] | None = None,
        json: Any = None,
    ) -> HttpResponse:
        """Serialize the optional JSON body and buffer the common response."""

        body = None
        if json is not None:
            body = jsonlib.dumps(
                json, separators=(",", ":"), ensure_ascii=False
            ).encode("utf-8")
        response = await self._transport.request(
            method,
            url,
            headers=headers,
            params=params,
            body=body,
            timeout=self._timeout_seconds,
        )
        return IntuneHttpResponseAdapter(
            response.status_code, response.headers, response.body
        )
