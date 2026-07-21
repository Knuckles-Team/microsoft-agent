"""Safe document payload and upload primitives for Microsoft Universal Print.

The Graph upload-session URL is preauthenticated.  It must never receive the
Graph bearer token, follow redirects, or be repurposed as an arbitrary network
target.  This module therefore accepts only the fixed documented Universal
Print upload origin and sends bounded sequential byte ranges.
"""

from __future__ import annotations

import asyncio
import base64
import binascii
import json
import re
from typing import Any
from urllib.parse import urlparse

from pydantic import BaseModel, ConfigDict, Field, field_validator

from microsoft_agent.power_platform import AsyncHttpTransport, HttpResponse

PRINT_UPLOAD_HOST = "print.print.microsoft.com"
PRINT_UPLOAD_ORIGIN = f"https://{PRINT_UPLOAD_HOST}"
PRINT_UPLOAD_CHUNK_BYTES = 16 * 320 * 1024
MAX_PRINT_DOCUMENT_BYTES = 100 * 1024 * 1024
_SAFE_CONTENT_TYPE = re.compile(
    r"^[A-Za-z0-9][A-Za-z0-9!#$&^_.+-]{0,126}/"
    r"[A-Za-z0-9][A-Za-z0-9!#$&^_.+-]{0,126}$"
)
_UPLOAD_PATH = re.compile(r"^/uploadSessions/[A-Za-z0-9._~-]{1,256}$")


class PrintDocumentSubmission(BaseModel):
    """Validated document and print configuration for one complete print job."""

    model_config = ConfigDict(extra="forbid")

    document_name: str = Field(min_length=1, max_length=255)
    content_type: str = Field(min_length=3, max_length=255)
    content_base64: str = Field(min_length=1, repr=False)
    configuration: dict[str, Any] = Field(default_factory=dict)

    @field_validator("document_name")
    @classmethod
    def validate_document_name(cls, value: str) -> str:
        """Reject paths, control characters, and ambiguous dot names."""

        name = value.strip()
        if (
            not name
            or name in {".", ".."}
            or "/" in name
            or "\\" in name
            or any(ord(char) < 32 for char in name)
        ):
            raise ValueError("document_name must be a plain safe filename")
        return name

    @field_validator("content_type")
    @classmethod
    def validate_content_type(cls, value: str) -> str:
        """Require a header-safe MIME media type without parameters."""

        media_type = value.strip().lower()
        if not _SAFE_CONTENT_TYPE.fullmatch(media_type):
            raise ValueError("content_type must be a safe MIME media type")
        return media_type

    @field_validator("configuration")
    @classmethod
    def validate_configuration(cls, value: dict[str, Any]) -> dict[str, Any]:
        """Keep the Graph printer configuration JSON bounded."""

        try:
            encoded = json.dumps(value, separators=(",", ":")).encode("utf-8")
        except (TypeError, ValueError) as exc:
            raise ValueError("configuration must be JSON serializable") from exc
        if len(encoded) > 64 * 1024:
            raise ValueError("configuration exceeds the 64 KiB safety limit")
        return value

    def content_bytes(self) -> bytes:
        """Decode the strict base64 document within the configured size bound."""

        if len(self.content_base64) > ((MAX_PRINT_DOCUMENT_BYTES + 2) // 3) * 4:
            raise ValueError("print document exceeds the 100 MiB safety limit")
        try:
            content = base64.b64decode(self.content_base64, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise ValueError("content_base64 must be valid base64") from exc
        if not content:
            raise ValueError("print document cannot be empty")
        if len(content) > MAX_PRINT_DOCUMENT_BYTES:
            raise ValueError("print document exceeds the 100 MiB safety limit")
        return content


class UniversalPrintUploadError(RuntimeError):
    """Safe upload failure that never contains the preauthenticated URL."""

    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        self.status_code = status_code
        super().__init__(message)


class UniversalPrintUploader:
    """Upload one print document to a Graph-created preauthenticated session."""

    def __init__(
        self,
        transport: AsyncHttpTransport,
        *,
        timeout_seconds: float = 60,
        chunk_bytes: int = PRINT_UPLOAD_CHUNK_BYTES,
        max_retries: int = 3,
    ) -> None:
        if not 0 < chunk_bytes < 10 * 1024 * 1024:
            raise ValueError("Universal Print chunks must be smaller than 10 MiB")
        if chunk_bytes % (320 * 1024):
            raise ValueError("Universal Print chunks must be multiples of 320 KiB")
        if not 0 < timeout_seconds <= 600:
            raise ValueError("timeout_seconds must be between 0 and 600")
        if not 0 <= max_retries <= 8:
            raise ValueError("max_retries must be between 0 and 8")
        self._transport = transport
        self.timeout_seconds = timeout_seconds
        self.chunk_bytes = chunk_bytes
        self.max_retries = max_retries

    async def upload(self, upload_url: str, content: bytes) -> dict[str, Any]:
        """Upload sequential ranges and return final printDocument metadata."""

        safe_url = validate_print_upload_url(upload_url)
        if not content:
            raise ValueError("print document cannot be empty")
        if len(content) > MAX_PRINT_DOCUMENT_BYTES:
            raise ValueError("print document exceeds the 100 MiB safety limit")

        total = len(content)
        offset = 0
        while offset < total:
            end = min(offset + self.chunk_bytes, total) - 1
            response = await self._put_range(
                safe_url,
                content[offset : end + 1],
                offset,
                end,
                total,
            )
            if response.status_code == 201:
                if end + 1 != total:
                    raise UniversalPrintUploadError(
                        "Universal Print completed before all bytes were uploaded."
                    )
                payload = _json_object(response)
                if not isinstance(payload.get("id"), str):
                    raise UniversalPrintUploadError(
                        "Universal Print returned invalid document metadata."
                    )
                return payload
            if response.status_code != 202:
                raise UniversalPrintUploadError(
                    "Universal Print rejected a document range.",
                    status_code=response.status_code,
                )
            payload = _json_object(response)
            ranges = payload.get("nextExpectedRanges")
            if not isinstance(ranges, list) or not ranges:
                raise UniversalPrintUploadError(
                    "Universal Print returned no next expected range."
                )
            next_offset = _range_start(ranges[0])
            if next_offset != end + 1 or next_offset >= total:
                raise UniversalPrintUploadError(
                    "Universal Print returned an unexpected next range."
                )
            offset = next_offset
        raise UniversalPrintUploadError(
            "Universal Print did not return completed document metadata."
        )

    async def _put_range(
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
        for attempt in range(self.max_retries + 1):
            try:
                response = await self._transport.request(
                    "PUT",
                    upload_url,
                    headers=headers,
                    body=fragment,
                    timeout=self.timeout_seconds,
                )
            except (OSError, TimeoutError) as exc:
                if attempt >= self.max_retries:
                    raise UniversalPrintUploadError(
                        "Universal Print upload transport failed."
                    ) from exc
                await asyncio.sleep(min(2**attempt, 8))
                continue
            if response.status_code in {201, 202}:
                return response
            if (
                response.status_code not in {429, 500, 502, 503, 504}
                or attempt >= self.max_retries
            ):
                return response
            await asyncio.sleep(_retry_delay(response, attempt))
        raise UniversalPrintUploadError("Universal Print retry limit was exceeded.")


def validate_print_upload_url(value: Any) -> str:
    """Allow only the opaque URL returned from the documented print origin."""

    if not isinstance(value, str) or not value or len(value) > 32768:
        raise UniversalPrintUploadError(
            "Universal Print returned an invalid upload URL."
        )
    parsed = urlparse(value)
    try:
        port = parsed.port
    except ValueError as exc:
        raise UniversalPrintUploadError(
            "Universal Print returned an invalid upload URL."
        ) from exc
    if (
        parsed.scheme != "https"
        or (parsed.hostname or "").casefold() != PRINT_UPLOAD_HOST
        or port not in {None, 443}
        or parsed.username
        or parsed.password
        or parsed.fragment
        or not _UPLOAD_PATH.fullmatch(parsed.path)
        or not parsed.query
    ):
        raise UniversalPrintUploadError(
            "Universal Print returned an unsafe upload URL."
        )
    return value


def _json_object(response: HttpResponse) -> dict[str, Any]:
    try:
        payload = response.json_body()
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise UniversalPrintUploadError(
            "Universal Print returned invalid JSON."
        ) from exc
    if not isinstance(payload, dict):
        raise UniversalPrintUploadError("Universal Print returned an invalid response.")
    return payload


def _range_start(value: Any) -> int:
    if not isinstance(value, str) or not re.fullmatch(r"\d+(?:-\d*)?", value):
        raise UniversalPrintUploadError(
            "Universal Print returned an invalid expected range."
        )
    return int(value.partition("-")[0])


def _retry_delay(response: HttpResponse, attempt: int) -> float:
    raw = next(
        (
            value
            for key, value in response.headers.items()
            if key.casefold() == "retry-after"
        ),
        None,
    )
    try:
        return (
            min(max(float(raw), 0.0), 30.0) if raw is not None else min(2**attempt, 8)
        )
    except ValueError:
        return min(2**attempt, 8)


__all__ = [
    "MAX_PRINT_DOCUMENT_BYTES",
    "PRINT_UPLOAD_CHUNK_BYTES",
    "PRINT_UPLOAD_HOST",
    "PrintDocumentSubmission",
    "UniversalPrintUploadError",
    "UniversalPrintUploader",
    "validate_print_upload_url",
]
