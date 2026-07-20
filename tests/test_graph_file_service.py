"""Tests for safe Graph file upload behavior."""

from __future__ import annotations

import json
from collections.abc import Mapping
from typing import Any

import pytest
from pydantic import ValidationError

from microsoft_agent.graph_file_service import (
    GRAPH_AUDIENCE,
    GraphFileService,
    GraphFileServiceError,
    GraphFileSettings,
    UploadConflictBehavior,
)
from microsoft_agent.power_platform import HttpResponse


class TokenProvider:
    def __init__(self) -> None:
        self.audiences: list[str] = []

    async def get_token(self, audience: str) -> str:
        self.audiences.append(audience)
        return "test-token"


class Transport:
    def __init__(self, responses: list[HttpResponse]) -> None:
        self.responses = responses
        self.requests: list[dict[str, Any]] = []

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
        self.requests.append(
            {
                "method": method,
                "url": url,
                "headers": dict(headers),
                "params": params,
                "body": body,
                "timeout": timeout,
            }
        )
        return self.responses.pop(0)


def response(status: int, payload: dict[str, Any]) -> HttpResponse:
    return HttpResponse(
        status_code=status,
        headers={"Content-Type": "application/json"},
        body=json.dumps(payload).encode(),
    )


@pytest.mark.asyncio
async def test_small_upload_uses_content_endpoint_and_graph_token() -> None:
    transport = Transport(
        [response(201, {"id": "item-1", "name": "Report.docx", "size": 3})]
    )
    tokens = TokenProvider()
    service = GraphFileService(GraphFileSettings(), tokens, transport)

    item = await service.upload_bytes(
        "drive-1",
        "Reports/Quarter 1.docx",
        b"doc",
        "application/test",
        conflict_behavior=UploadConflictBehavior.RENAME,
    )

    assert item.item_id == "item-1"
    assert tokens.audiences == [GRAPH_AUDIENCE]
    request = transport.requests[0]
    assert request["method"] == "PUT"
    assert "Quarter%201.docx" in request["url"]
    assert request["headers"]["Authorization"] == "Bearer test-token"
    assert request["params"] == {"@microsoft.graph.conflictBehavior": "rename"}


@pytest.mark.asyncio
async def test_large_upload_uses_sequential_preauthenticated_fragments() -> None:
    chunk = 320 * 1024
    content = b"a" * (chunk + 1)
    transport = Transport(
        [
            response(200, {"uploadUrl": "https://upload.example.test/session"}),
            response(202, {"nextExpectedRanges": [f"{chunk}-"]}),
            response(
                201,
                {"id": "item-2", "name": "deck.pptx", "size": len(content)},
            ),
        ]
    )
    service = GraphFileService(
        GraphFileSettings(
            resumable_threshold_bytes=1,
            fragment_size_bytes=chunk,
        ),
        TokenProvider(),
        transport,
    )

    item = await service.upload_bytes("drive-1", "deck.pptx", content)

    assert item.item_id == "item-2"
    assert transport.requests[0]["method"] == "POST"
    fragments = transport.requests[1:]
    assert [request["headers"]["Content-Range"] for request in fragments] == [
        f"bytes 0-{chunk - 1}/{len(content)}",
        f"bytes {chunk}-{chunk}/{len(content)}",
    ]
    assert all("Authorization" not in request["headers"] for request in fragments)


@pytest.mark.parametrize(
    "upload_url",
    [
        "https://127.0.0.1/session",
        "https://10.0.0.1/session",
        "https://169.254.169.254/latest/meta-data",
        "https://[::1]/session",
        "https://upload.example.test:8443/session",
        "https://user:password@upload.example.test/session",
        "https://upload.example.test/session#fragment",
        "https://service.local/session",
        " https://upload.example.test/session",
        r"https://upload.example.test\@127.0.0.1/session",
    ],
)
@pytest.mark.asyncio
async def test_resumable_upload_rejects_unsafe_response_targets(
    upload_url: str,
) -> None:
    transport = Transport([response(200, {"uploadUrl": upload_url})])
    service = GraphFileService(
        GraphFileSettings(resumable_threshold_bytes=1),
        TokenProvider(),
        transport,
    )

    with pytest.raises(GraphFileServiceError, match="unsafe upload URL"):
        await service.upload_bytes("drive-1", "report.bin", b"x")

    assert len(transport.requests) == 1


@pytest.mark.parametrize(
    "path",
    ["", "/absolute.docx", "../escape.docx", "folder\\file.docx", "bad:name.docx"],
)
@pytest.mark.asyncio
async def test_upload_rejects_unsafe_drive_paths(path: str) -> None:
    service = GraphFileService(GraphFileSettings(), TokenProvider(), Transport([]))
    with pytest.raises(ValueError):
        await service.upload_bytes("drive-1", path, b"x")


def test_fragment_size_must_be_graph_granularity() -> None:
    with pytest.raises(ValidationError, match="320 KiB"):
        GraphFileSettings(fragment_size_bytes=400_000)


def test_graph_base_url_cannot_redirect_tokens() -> None:
    with pytest.raises(ValidationError, match="graph.microsoft.com"):
        GraphFileSettings(graph_base_url="https://example.test/v1.0")
