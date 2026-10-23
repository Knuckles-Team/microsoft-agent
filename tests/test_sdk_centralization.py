"""TLS profile and upload retry behavior owned by the connector SDK."""

from __future__ import annotations

import httpx
import pytest
from agent_connector_sdk.http.client import create_http_client
from agent_connector_sdk.http.options import HttpClientOptions
from agent_connector_sdk.http.retry import RetryPolicy

from microsoft_agent.graph_file_service import (
    GraphFileService,
    GraphFileServiceError,
    GraphFileSettings,
)
from microsoft_agent.power_platform import (
    HttpxAsyncHttpTransport,
    upload_retry_policy,
)
from microsoft_agent.universal_print import UniversalPrintUploader

PRINT_URL = (
    "https://print.print.microsoft.com/uploadSessions/session-1?tempauthtoken=opaque"
)
UPLOAD_URL = "https://upload.example/session-1?tempauthtoken=opaque"


class _NoToken:
    async def get_token(self, audience: str) -> str:
        return "token"


def _transport(
    handler: httpx.MockTransport, retries: int, monkeypatch: pytest.MonkeyPatch
) -> HttpxAsyncHttpTransport:
    monkeypatch.setattr("agent_connector_sdk.http.transport.time.sleep", lambda _: None)
    client = create_http_client(
        HttpClientOptions(
            base_url="https://pinned.invalid",
            retry=upload_retry_policy(retries),
        ),
        transport=handler,
    )
    return HttpxAsyncHttpTransport(service="microsoft_graph", client=client)


@pytest.mark.spec("CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001")
def test_upload_retry_policy_is_bounded_and_retries_server_errors() -> None:
    policy = upload_retry_policy(3)
    assert policy.max_attempts == 4
    assert {429, 500, 503} <= policy.retry_statuses
    assert policy.max_retry_after == 30.0
    assert isinstance(policy, RetryPolicy)


@pytest.mark.spec("CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001")
async def test_print_range_retries_through_the_sdk_client(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[int] = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(1)
        if len(calls) < 3:
            return httpx.Response(503, headers={"Retry-After": "1"})
        return httpx.Response(201, json={"id": "document-1"})

    uploader = UniversalPrintUploader(
        _transport(httpx.MockTransport(handler), 3, monkeypatch)
    )
    result = await uploader.upload(PRINT_URL, b"abc")
    assert result == {"id": "document-1"}
    assert len(calls) == 3


@pytest.mark.spec("CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001")
async def test_graph_fragment_failure_surfaces_status_after_sdk_retries(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[int] = []

    def handler(request: httpx.Request) -> httpx.Response:
        calls.append(1)
        return httpx.Response(500)

    transport = _transport(httpx.MockTransport(handler), 2, monkeypatch)
    service = GraphFileService(
        GraphFileSettings(max_retries=2),
        _NoToken(),
        transport,  # type: ignore[arg-type]
    )
    with pytest.raises(GraphFileServiceError) as caught:
        await service._upload_fragment(UPLOAD_URL, b"abc", 0, 2, 3)  # noqa: SLF001
    assert caught.value.status_code == 500
    assert len(calls) == 3
