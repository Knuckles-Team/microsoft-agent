"""Safe OneDrive and SharePoint document upload support.

Small files use the Microsoft Graph content endpoint.  Files over the
configured threshold use a resumable upload session with sequential chunks
that are multiples of 320 KiB, as required by Graph.
"""

from __future__ import annotations

import asyncio
import json
import re
from collections.abc import Mapping
from enum import StrEnum
from pathlib import PurePosixPath
from typing import Any
from urllib.parse import quote, urlparse

from agent_utilities.security.egress import validate_base_url
from pydantic import BaseModel, ConfigDict, Field, HttpUrl, field_validator

from microsoft_agent.document_service import GeneratedArtifact
from microsoft_agent.power_platform import (
    AsyncHttpTransport,
    AudienceTokenProvider,
    HttpResponse,
)

GRAPH_AUDIENCE = "https://graph.microsoft.com"
_CHUNK_GRANULARITY = 320 * 1024
_MAX_FRAGMENT_BYTES = 60 * 1024 * 1024
_SAFE_DRIVE_ID = re.compile(r"^[A-Za-z0-9!._~-]{1,512}$")


class UploadConflictBehavior(StrEnum):
    """How Graph should handle an existing destination name."""

    FAIL = "fail"
    REPLACE = "replace"
    RENAME = "rename"


class GraphFileSettings(BaseModel):
    """Validated Microsoft Graph upload configuration."""

    model_config = ConfigDict(frozen=True)

    graph_base_url: HttpUrl = HttpUrl("https://graph.microsoft.com/v1.0")
    timeout_seconds: float = Field(default=60.0, gt=0, le=600)
    resumable_threshold_bytes: int = Field(
        default=10 * 1024 * 1024, ge=1, le=250 * 1024 * 1024
    )
    fragment_size_bytes: int = Field(
        default=10 * _CHUNK_GRANULARITY,
        ge=_CHUNK_GRANULARITY,
        lt=_MAX_FRAGMENT_BYTES,
    )
    max_file_bytes: int = Field(default=250 * 1024 * 1024, ge=1)
    max_retries: int = Field(default=3, ge=0, le=8)

    @field_validator("graph_base_url")
    @classmethod
    def validate_graph_base_url(cls, value: HttpUrl) -> HttpUrl:
        """Require the official HTTPS Graph host and a stable API version."""

        if (
            value.scheme != "https"
            or (value.host or "").lower() != "graph.microsoft.com"
        ):
            raise ValueError("graph_base_url must use https://graph.microsoft.com")
        if (value.path or "").rstrip("/") != "/v1.0" or value.query or value.fragment:
            raise ValueError("graph_base_url must target Microsoft Graph v1.0")
        return value

    @field_validator("fragment_size_bytes")
    @classmethod
    def validate_fragment_size(cls, value: int) -> int:
        """Graph fragments must be exact multiples of 320 KiB."""

        if value % _CHUNK_GRANULARITY:
            raise ValueError("fragment_size_bytes must be a multiple of 320 KiB")
        return value


class UploadedDriveItem(BaseModel):
    """Normalized metadata returned for a completed Graph upload."""

    model_config = ConfigDict(populate_by_name=True, extra="allow")

    item_id: str = Field(alias="id")
    name: str
    size: int = Field(ge=0)
    web_url: str | None = Field(default=None, alias="webUrl")
    e_tag: str | None = Field(default=None, alias="eTag")
    c_tag: str | None = Field(default=None, alias="cTag")


class GraphFileServiceError(RuntimeError):
    """Safe Graph upload failure without token or upload-URL disclosure."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        self.status_code = status_code
        super().__init__(message)


class GraphFileService:
    """Upload generated or arbitrary bytes to a configured Microsoft drive."""

    def __init__(
        self,
        settings: GraphFileSettings,
        token_provider: AudienceTokenProvider,
        transport: AsyncHttpTransport,
    ) -> None:
        self.settings = settings
        self._token_provider = token_provider
        self._transport = transport

    async def upload_artifact(
        self,
        drive_id: str,
        destination_path: str,
        artifact: GeneratedArtifact,
        *,
        conflict_behavior: UploadConflictBehavior = UploadConflictBehavior.FAIL,
    ) -> UploadedDriveItem:
        """Upload an in-memory generated Office artifact."""

        return await self.upload_bytes(
            drive_id,
            destination_path,
            artifact.upload_bytes(),
            artifact.content_type,
            conflict_behavior=conflict_behavior,
        )

    async def upload_bytes(
        self,
        drive_id: str,
        destination_path: str,
        content: bytes,
        content_type: str = "application/octet-stream",
        *,
        conflict_behavior: UploadConflictBehavior = UploadConflictBehavior.FAIL,
    ) -> UploadedDriveItem:
        """Upload bytes by drive-relative path, using a session when needed."""

        encoded_drive = _encode_drive_id(drive_id)
        encoded_path = _encode_drive_path(destination_path)
        if not content:
            raise ValueError("content cannot be empty")
        if len(content) > self.settings.max_file_bytes:
            raise ValueError("content exceeds the configured upload size limit")
        if not content_type or "\r" in content_type or "\n" in content_type:
            raise ValueError("content_type must be a safe MIME type")
        if len(content) < self.settings.resumable_threshold_bytes:
            return await self._simple_upload(
                encoded_drive,
                encoded_path,
                content,
                content_type,
                conflict_behavior,
            )
        return await self._resumable_upload(
            encoded_drive, encoded_path, content, conflict_behavior
        )

    async def _simple_upload(
        self,
        drive_id: str,
        path: str,
        content: bytes,
        content_type: str,
        conflict_behavior: UploadConflictBehavior,
    ) -> UploadedDriveItem:
        base = str(self.settings.graph_base_url).rstrip("/")
        url = f"{base}/drives/{drive_id}/root:/{path}:/content"
        response = await self._graph_request(
            "PUT",
            url,
            headers={"Accept": "application/json", "Content-Type": content_type},
            params={"@microsoft.graph.conflictBehavior": conflict_behavior.value},
            body=content,
            expected={200, 201},
        )
        return self._drive_item(response)

    async def _resumable_upload(
        self,
        drive_id: str,
        path: str,
        content: bytes,
        conflict_behavior: UploadConflictBehavior,
    ) -> UploadedDriveItem:
        base = str(self.settings.graph_base_url).rstrip("/")
        filename = PurePosixPath(path).name
        session_url = f"{base}/drives/{drive_id}/root:/{path}:/createUploadSession"
        session_body = json.dumps(
            {
                "item": {
                    "@microsoft.graph.conflictBehavior": conflict_behavior.value,
                    "name": filename,
                }
            },
            separators=(",", ":"),
        ).encode("utf-8")
        response = await self._graph_request(
            "POST",
            session_url,
            headers={"Accept": "application/json", "Content-Type": "application/json"},
            body=session_body,
            expected={200, 201},
        )
        payload = _json_object(response)
        upload_url = _validate_upload_url(payload.get("uploadUrl"))

        total = len(content)
        offset = 0
        while offset < total:
            end = min(offset + self.settings.fragment_size_bytes, total) - 1
            fragment = content[offset : end + 1]
            chunk_response = await self._upload_fragment(
                upload_url,
                fragment,
                offset,
                end,
                total,
            )
            if chunk_response.status_code in {200, 201}:
                if end + 1 != total:
                    raise GraphFileServiceError(
                        "Graph completed the upload before all bytes were sent."
                    )
                return self._drive_item(chunk_response)
            payload = _json_object(chunk_response)
            ranges = payload.get("nextExpectedRanges")
            if not isinstance(ranges, list) or not ranges:
                raise GraphFileServiceError(
                    "Graph returned no next range for the resumable upload."
                )
            next_offset = _range_start(ranges[0])
            if next_offset < end + 1 or next_offset > total:
                raise GraphFileServiceError(
                    "Graph returned an invalid next range for the upload."
                )
            offset = next_offset
        raise GraphFileServiceError(
            "Graph did not return completed drive item metadata."
        )

    async def _upload_fragment(
        self,
        upload_url: str,
        fragment: bytes,
        start: int,
        end: int,
        total: int,
    ) -> HttpResponse:
        headers = {
            "Content-Length": str(len(fragment)),
            "Content-Range": f"bytes {start}-{end}/{total}",
        }
        # The preauthenticated upload URL must not receive the Graph bearer token.
        for attempt in range(self.settings.max_retries + 1):
            try:
                response = await self._transport.request(
                    "PUT",
                    upload_url,
                    headers=headers,
                    body=fragment,
                    timeout=self.settings.timeout_seconds,
                )
            except (OSError, TimeoutError) as exc:
                if attempt >= self.settings.max_retries:
                    raise GraphFileServiceError("Upload transport failed.") from exc
                await asyncio.sleep(min(2**attempt, 8))
                continue
            if response.status_code in {200, 201, 202}:
                return response
            if (
                response.status_code not in {429, 500, 502, 503, 504}
                or attempt >= self.settings.max_retries
            ):
                raise GraphFileServiceError(
                    "Microsoft Graph rejected an upload fragment.",
                    status_code=response.status_code,
                )
            await asyncio.sleep(_retry_delay(response, attempt))
        raise GraphFileServiceError("Upload fragment retry limit was exceeded.")

    async def _graph_request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        body: bytes | None = None,
        params: Mapping[str, Any] | None = None,
        expected: set[int],
    ) -> HttpResponse:
        token = await self._token_provider.get_token(GRAPH_AUDIENCE)
        request_headers = {**dict(headers), "Authorization": f"Bearer {token}"}
        try:
            response = await self._transport.request(
                method,
                url,
                headers=request_headers,
                params=params,
                body=body,
                timeout=self.settings.timeout_seconds,
            )
        except (OSError, TimeoutError) as exc:
            raise GraphFileServiceError("Microsoft Graph transport failed.") from exc
        if response.status_code not in expected:
            raise GraphFileServiceError(
                "Microsoft Graph rejected the file operation.",
                status_code=response.status_code,
            )
        return response

    @staticmethod
    def _drive_item(response: HttpResponse) -> UploadedDriveItem:
        try:
            return UploadedDriveItem.model_validate(_json_object(response))
        except (TypeError, ValueError) as exc:
            raise GraphFileServiceError(
                "Microsoft Graph returned invalid drive item metadata."
            ) from exc


def _encode_drive_id(drive_id: str) -> str:
    value = drive_id.strip()
    if not _SAFE_DRIVE_ID.fullmatch(value):
        raise ValueError("drive_id contains unsupported characters")
    return quote(value, safe="!._~-")


def _encode_drive_path(path: str) -> str:
    if not path or path.startswith("/") or "\\" in path or "\x00" in path:
        raise ValueError("destination_path must be a relative drive path")
    parsed = PurePosixPath(path)
    if any(part in {"", ".", ".."} for part in parsed.parts):
        raise ValueError("destination_path cannot contain dot segments")
    if any(
        any(ord(char) < 32 for char in part) or ":" in part for part in parsed.parts
    ):
        raise ValueError("destination_path contains unsupported characters")
    if len(path) > 400:
        raise ValueError("destination_path is too long")
    return "/".join(quote(part, safe="!$&'()+,;=@[]^_`{}~-") for part in parsed.parts)


def _validate_upload_url(value: Any) -> str:
    if not isinstance(value, str) or not value or len(value) > 8_192:
        raise GraphFileServiceError("Graph did not return a valid upload URL.")
    if (
        value != value.strip()
        or "\\" in value
        or any(ord(character) < 32 for character in value)
    ):
        raise GraphFileServiceError("Graph returned an unsafe upload URL.")
    parsed = urlparse(value)
    try:
        port = parsed.port
        host = (parsed.hostname or "").casefold().rstrip(".")
        host.encode("ascii")
    except (UnicodeEncodeError, ValueError) as exc:
        raise GraphFileServiceError("Graph returned an unsafe upload URL.") from exc
    decision = validate_base_url(value, allow_loopback=False)
    if (
        not decision.allowed
        or parsed.scheme != "https"
        or not host
        or port not in {None, 443}
        or parsed.username
        or parsed.password
        or parsed.fragment
        or not parsed.path
        or host == "localhost"
        or host.endswith((".localhost", ".local", ".internal", ".home.arpa"))
    ):
        raise GraphFileServiceError("Graph returned an unsafe upload URL.")
    return value


def _json_object(response: HttpResponse) -> dict[str, Any]:
    try:
        value = response.json_body()
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise GraphFileServiceError("Microsoft Graph returned invalid JSON.") from exc
    if not isinstance(value, dict):
        raise GraphFileServiceError("Microsoft Graph returned an invalid response.")
    return value


def _range_start(value: Any) -> int:
    if not isinstance(value, str) or not re.fullmatch(r"\d+-\d*", value):
        raise GraphFileServiceError("Graph returned an invalid expected range.")
    return int(value.partition("-")[0])


def _retry_delay(response: HttpResponse, attempt: int) -> float:
    raw = next(
        (
            value
            for key, value in response.headers.items()
            if key.lower() == "retry-after"
        ),
        None,
    )
    try:
        return (
            min(max(float(raw), 0.0), 30.0) if raw is not None else min(2**attempt, 8)
        )
    except ValueError:
        return min(2**attempt, 8)
