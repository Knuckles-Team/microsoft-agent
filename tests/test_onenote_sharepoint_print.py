"""Regression tests for OneNote, SharePoint delta, and Universal Print paths."""

from __future__ import annotations

import base64
import json
from collections.abc import Mapping
from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from fastmcp import FastMCP
from kiota_abstractions.method import Method
from pydantic import ValidationError

from microsoft_agent.api_client import MicrosoftGraphApi
from microsoft_agent.power_platform import HttpResponse
from microsoft_agent.universal_print import (
    PRINT_UPLOAD_CHUNK_BYTES,
    PrintDocumentSubmission,
    UniversalPrintUploader,
    UniversalPrintUploadError,
    validate_print_upload_url,
)


class RecordingTransport:
    """Small no-network transport that records preauthenticated requests."""

    def __init__(self, responses: list[HttpResponse]) -> None:
        self.responses = list(responses)
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
        body=json.dumps(payload).encode("utf-8"),
    )


def authenticated_api(
    mock_auth_manager: MagicMock, client: MagicMock
) -> MicrosoftGraphApi:
    mock_auth_manager.get_token.return_value = "test-token"
    mock_auth_manager.get_current_account.return_value = {
        "username": "test@example.com"
    }
    with patch(
        "microsoft_agent.api.api_client_base.GraphServiceClient", return_value=client
    ):
        return MicrosoftGraphApi(mock_auth_manager)


@pytest.mark.asyncio
async def test_onenote_create_sends_raw_html_request(mock_auth_manager) -> None:
    native = MagicMock()
    native.json.return_value = {"id": "page-1"}
    pages = MagicMock()
    pages.url_template = "{+baseurl}/me/onenote/pages"
    pages.path_parameters = {"baseurl": "https://graph.microsoft.com/v1.0"}
    pages.to_post_request_configuration.return_value = MagicMock(options=[])
    pages.request_adapter.send_async = AsyncMock(return_value=native)
    client = MagicMock()
    client.me.onenote.pages = pages
    api = authenticated_api(mock_auth_manager, client)
    html = "<!doctype html><html><head><title>Plan</title></head><body>OK</body></html>"

    result = await api.create_onenote_page({"content": html})

    assert result == {"id": "page-1"}
    awaited_send = pages.request_adapter.send_async.await_args
    assert awaited_send is not None
    request_info = awaited_send.args[0]
    assert request_info.http_method is Method.POST
    assert request_info.content == html.encode("utf-8")
    assert request_info.headers.get("Content-Type") == {"text/html"}


@pytest.mark.asyncio
async def test_sharepoint_site_by_path_uses_current_sdk_builder(
    mock_auth_manager, mock_native_response
) -> None:
    client = MagicMock()
    site = client.sites.by_site_id.return_value
    by_path = site.get_by_path_with_path.return_value
    by_path.to_get_request_configuration.return_value = MagicMock(options=[])
    by_path.get = AsyncMock(return_value=mock_native_response)
    api = authenticated_api(mock_auth_manager, client)

    result = await api.get_sharepoint_site_by_path("contoso.sharepoint.com", "/teams/a")

    assert result == {"value": []}
    site.get_by_path_with_path.assert_called_with("/teams/a")
    assert not site.get_by_path.called


@pytest.mark.asyncio
async def test_sharepoint_delta_can_resume_and_exhaust_pages(mock_auth_manager) -> None:
    next_url = "https://graph.microsoft.com/v1.0/sites/delta?token=next"
    delta_url = "https://graph.microsoft.com/v1.0/sites/delta?$deltatoken=done"
    first_response = MagicMock()
    first_response.json.return_value = {
        "value": [{"id": "site-1"}],
        "@odata.nextLink": next_url,
    }
    second_response = MagicMock()
    second_response.json.return_value = {
        "value": [{"id": "site-2"}],
        "@odata.deltaLink": delta_url,
    }
    first = MagicMock()
    first.get = AsyncMock(return_value=first_response)
    second = MagicMock()
    second.to_get_request_configuration.return_value = MagicMock(options=[])
    second.get = AsyncMock(return_value=second_response)
    client = MagicMock()
    client.sites.delta.with_url.side_effect = [first, second]
    api = authenticated_api(mock_auth_manager, client)

    result = await api.get_sharepoint_sites_delta(
        params={"token": "latest", "$top": "2"}, fetch_all=True
    )

    assert result == {
        "value": [{"id": "site-1"}, {"id": "site-2"}],
        "pagesFetched": 2,
        "@odata.deltaLink": delta_url,
    }
    first_url = client.sites.delta.with_url.call_args_list[0].args[0]
    assert first_url.startswith("https://graph.microsoft.com/v1.0/sites/delta?")
    assert "token=latest" in first_url
    assert client.sites.delta.with_url.call_args_list[1].args[0] == next_url


@pytest.mark.asyncio
async def test_sharepoint_delta_rejects_cross_origin_continuation(
    mock_auth_manager,
) -> None:
    client = MagicMock()
    api = authenticated_api(mock_auth_manager, client)

    result = await api.get_sharepoint_sites_delta(
        continuation_url="https://evil.example/sites/delta?token=secret"
    )

    assert "Microsoft Graph v1.0 sites/delta" in result["error"]
    client.sites.delta.with_url.assert_not_called()


@pytest.mark.asyncio
async def test_universal_print_uploads_bounded_ranges_without_bearer() -> None:
    content = b"a" * (PRINT_UPLOAD_CHUNK_BYTES + 3)
    upload_url = (
        "https://print.print.microsoft.com/uploadSessions/session-1"
        "?tempauthtoken=opaque"
    )
    transport = RecordingTransport(
        [
            response(
                202,
                {
                    "nextExpectedRanges": [
                        f"{PRINT_UPLOAD_CHUNK_BYTES}-{len(content) - 1}"
                    ]
                },
            ),
            response(
                201,
                {
                    "id": "document-1",
                    "documentName": "report.pdf",
                    "contentType": "application/pdf",
                    "size": len(content),
                },
            ),
        ]
    )

    result = await UniversalPrintUploader(transport).upload(upload_url, content)

    assert result["id"] == "document-1"
    assert [item["headers"]["Content-Range"] for item in transport.requests] == [
        f"bytes 0-{PRINT_UPLOAD_CHUNK_BYTES - 1}/{len(content)}",
        f"bytes {PRINT_UPLOAD_CHUNK_BYTES}-{len(content) - 1}/{len(content)}",
    ]
    assert all("Authorization" not in item["headers"] for item in transport.requests)
    assert all(item["url"] == upload_url for item in transport.requests)


@pytest.mark.parametrize(
    "url",
    [
        "http://print.print.microsoft.com/uploadSessions/a?tempauthtoken=x",
        "https://evil.example/uploadSessions/a?tempauthtoken=x",
        "https://print.print.microsoft.com.evil.example/uploadSessions/a?x=1",
        "https://print.print.microsoft.com/other/a?tempauthtoken=x",
    ],
)
def test_universal_print_rejects_unsafe_upload_origins(url: str) -> None:
    with pytest.raises(UniversalPrintUploadError, match="unsafe upload URL"):
        validate_print_upload_url(url)


@pytest.mark.asyncio
async def test_print_graph_methods_use_real_sdk_root_and_typed_bodies(
    mock_auth_manager,
) -> None:
    client = MagicMock()
    jobs = client.print.printers.by_printer_id.return_value.jobs
    job_response = MagicMock()
    job_response.json.return_value = {"id": "job-1"}
    jobs.post = AsyncMock(return_value=job_response)
    session_builder = jobs.by_print_job_id.return_value.documents.by_print_document_id.return_value.create_upload_session
    session_builder.to_post_request_configuration.return_value = MagicMock(options=[])
    session_response = MagicMock()
    session_response.json.return_value = {"uploadUrl": "opaque"}
    session_builder.post = AsyncMock(return_value=session_response)
    start = jobs.by_print_job_id.return_value.start
    start.to_post_request_configuration.return_value = MagicMock(options=[])
    start_response = MagicMock()
    start_response.json.return_value = {"state": "processing"}
    start.post = AsyncMock(return_value=start_response)
    api = authenticated_api(mock_auth_manager, client)

    await api.create_print_job("printer-1", {"configuration": {"copies": 2}})
    await api.create_print_document_upload_session(
        "printer-1", "job-1", "document-1", "report.pdf", "application/pdf", 12
    )
    await api.start_print_job("printer-1", "job-1")

    awaited_job = jobs.post.await_args
    assert awaited_job is not None
    job_body = awaited_job.args[0]
    assert job_body.configuration.copies == 2
    awaited_session = session_builder.post.await_args
    assert awaited_session is not None
    session_body = awaited_session.args[0]
    assert session_body.properties.document_name == "report.pdf"
    assert session_body.properties.content_type == "application/pdf"
    assert session_body.properties.size == 12
    start.post.assert_awaited_once()
    assert not hasattr(client, "print_") or not client.print_.mock_calls


@pytest.mark.asyncio
async def test_submit_print_document_runs_complete_lifecycle(mock_auth_manager) -> None:
    api = authenticated_api(mock_auth_manager, MagicMock())
    create_job = AsyncMock(
        return_value={"id": "job-1", "documents": [{"id": "document-1"}]}
    )
    upload_url = (
        "https://print.print.microsoft.com/uploadSessions/session-1"
        "?tempauthtoken=opaque"
    )
    create_session = AsyncMock(return_value={"uploadUrl": upload_url})
    start_job = AsyncMock(return_value={"state": "processing"})
    transport = RecordingTransport(
        [
            response(
                201,
                {
                    "id": "document-1",
                    "documentName": "report.pdf",
                    "contentType": "application/pdf",
                    "size": 3,
                },
            )
        ]
    )
    submission = PrintDocumentSubmission(
        document_name="report.pdf",
        content_type="application/pdf",
        content_base64=base64.b64encode(b"pdf").decode("ascii"),
        configuration={"copies": 2},
    )

    with (
        patch.object(api, "create_print_job", create_job),
        patch.object(api, "create_print_document_upload_session", create_session),
        patch.object(api, "start_print_job", start_job),
    ):
        result = await api.submit_print_document(
            "printer-1", submission, upload_transport=transport
        )

    assert result["jobId"] == "job-1"
    assert result["documentId"] == "document-1"
    assert result["status"] == {"state": "processing"}
    create_job.assert_awaited_once_with("printer-1", {"configuration": {"copies": 2}})
    start_job.assert_awaited_once_with("printer-1", "job-1")
    assert "Authorization" not in transport.requests[0]["headers"]


def test_print_submission_validates_filename_and_base64() -> None:
    with pytest.raises(ValidationError, match="plain safe filename"):
        PrintDocumentSubmission(
            document_name="../report.pdf",
            content_type="application/pdf",
            content_base64="cGRm",
        )
    request = PrintDocumentSubmission(
        document_name="report.pdf",
        content_type="application/pdf",
        content_base64="not base64!",
    )
    with pytest.raises(ValueError, match="valid base64"):
        request.content_bytes()


@pytest.mark.asyncio
async def test_print_tools_register_typed_submit_operation() -> None:
    from microsoft_agent.mcp_server import register_print_tools

    mcp = FastMCP("print-test")
    register_print_tools(mcp)

    names = {tool.name for tool in await mcp.list_tools()}
    assert names == {"microsoft_print"}
    tool = await mcp.get_tool("microsoft_print")
    actions = await tool.fn(
        action="list_actions",
        params_json="{}",
        client=MagicMock(),
        ctx=None,
    )
    assert "submit_print_document" in actions["actions"]
