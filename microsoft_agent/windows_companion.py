"""Typed control-plane client for a native Windows companion.

The companion is modeled as an agent that establishes an authenticated,
outbound connection to an HTTPS control plane.  This client talks only to that
control plane; it never opens a direct connection to a laptop.  Devices,
actions, files, Windows services, and Power Automate Desktop flows are all
explicitly allowlisted.

There is intentionally no shell, command-line, PowerShell, arbitrary process,
or arbitrary URL action.  Adding a new capability requires a new typed action
model and an explicit policy decision in both controller and companion.
"""

from __future__ import annotations

import asyncio
import base64
import binascii
import json
import ntpath
import re
from collections.abc import Mapping
from datetime import UTC, datetime, timedelta
from enum import StrEnum
from typing import Annotated, Any, Literal, Never, Protocol, runtime_checkable
from urllib.parse import quote
from uuid import UUID, uuid4

import httpx
from agent_utilities.core.http_client import create_http_client
from agent_utilities.core.transport_security import (
    ResolvedTLSProfile,
    resolve_configured_tls_profile,
)
from agent_utilities.security.cli_secrets import validate_runtime_secret_reference
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    HttpUrl,
    field_validator,
    model_validator,
)

from microsoft_agent.integration_auth import normalize_audience


class CompanionHttpResponse(BaseModel):
    """Transport-neutral HTTP response from the companion control plane."""

    model_config = ConfigDict(frozen=True)

    status_code: int = Field(ge=100, le=599)
    headers: dict[str, str] = Field(default_factory=dict)
    body: bytes = b""

    def json_body(self) -> Any:
        """Decode a JSON response body."""

        if not self.body:
            return None
        return json.loads(self.body.decode("utf-8"))


@runtime_checkable
class CompanionHttpTransport(Protocol):
    """Minimal injectable asynchronous HTTP transport."""

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        timeout: float,
    ) -> CompanionHttpResponse:
        """Send one request without automatically following redirects."""


@runtime_checkable
class CompanionTokenProvider(Protocol):
    """Acquire an access token for the companion control-plane audience."""

    async def get_token(self, audience: str) -> str:
        """Return a bearer token for ``audience``."""


class HttpxCompanionTransport:
    """Bounded async facade over the shared verified sync HTTP boundary."""

    def __init__(
        self,
        *,
        tls_profile: str | None = None,
        tls_profile_ref: str | None = None,
        allowed_private_hosts: tuple[str, ...] = (),
        max_response_bytes: int = 16 * 1024 * 1024,
        client: httpx.Client | None = None,
    ) -> None:
        if not 1_024 <= max_response_bytes <= 64 * 1024 * 1024:
            raise ValueError("HTTP response bound is invalid")
        self._max_response_bytes = max_response_bytes
        self._tls: ResolvedTLSProfile | None = None
        self._owns_client = client is None
        self._closed = False
        if client is not None:
            self._client = client
            return
        self._tls = resolve_configured_tls_profile(
            "microsoft_companion",
            profile_name=tls_profile,
            profile_ref=tls_profile_ref,
        )
        if self._tls.proxy_url:
            self._tls.cleanup()
            self._tls = None
            raise ValueError("Pinned provider transport does not support a proxy")
        try:
            self._client = create_http_client(
                timeout=httpx.Timeout(30.0),
                verify=self._tls.ssl_context,
                follow_redirects=False,
                trust_env=False,
                pin_egress=True,
                allowed_private_hosts=allowed_private_hosts,
                limits=httpx.Limits(
                    max_connections=32,
                    max_keepalive_connections=8,
                ),
            )
        except Exception:
            self._tls.cleanup()
            self._tls = None
            raise

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        timeout: float,
    ) -> CompanionHttpResponse:
        """Send one request without following redirects."""

        try:
            response = await asyncio.to_thread(
                self._client.request,
                method,
                url,
                headers=dict(headers),
                params=params,
                content=body,
                timeout=timeout,
                follow_redirects=False,
            )
        except httpx.TimeoutException:
            raise TimeoutError("Companion request timed out") from None
        except httpx.TransportError:
            raise OSError("Companion transport failed") from None
        if len(response.content) > self._max_response_bytes:
            raise ValueError("Provider response exceeds the configured bound")
        return CompanionHttpResponse(
            status_code=response.status_code,
            headers=dict(response.headers),
            body=response.content,
        )

    def close(self) -> None:
        """Close owned HTTP state and remove materialized trust files."""
        if self._closed:
            return
        self._closed = True
        try:
            if self._owns_client:
                self._client.close()
        finally:
            if self._tls is not None:
                self._tls.cleanup()
                self._tls = None


class CompanionActionKind(StrEnum):
    """Closed set of supported actions; arbitrary execution is absent."""

    SYSTEM_INVENTORY = "system.inventory"
    FILE_LIST = "file.list"
    FILE_READ = "file.read"
    FILE_WRITE = "file.write"
    OFFICE_OPEN_DOCUMENT = "office.open_document"
    OFFICE_EXPORT_PDF = "office.export_pdf"
    POWER_AUTOMATE_DESKTOP_RUN = "power_automate_desktop.run"
    WINDOWS_SERVICE_STATUS = "windows.service.status"
    WINDOWS_SERVICE_START = "windows.service.start"
    WINDOWS_SERVICE_STOP = "windows.service.stop"
    NOTIFICATION_SHOW = "notification.show"
    CLIPBOARD_READ_TEXT = "clipboard.read_text"
    CLIPBOARD_WRITE_TEXT = "clipboard.write_text"


class _ActionModel(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)


class SystemInventoryAction(_ActionModel):
    """Request bounded system, software, and device inventory."""

    kind: Literal[CompanionActionKind.SYSTEM_INVENTORY] = (
        CompanionActionKind.SYSTEM_INVENTORY
    )
    include_software: bool = False
    include_network_adapters: bool = False


class FileListAction(_ActionModel):
    """List entries below an allowlisted Windows path."""

    kind: Literal[CompanionActionKind.FILE_LIST] = CompanionActionKind.FILE_LIST
    path: str = Field(min_length=3, max_length=1024)
    recursive: bool = False
    max_entries: int = Field(default=200, ge=1, le=2000)


class FileReadAction(_ActionModel):
    """Read a bounded file below an allowlisted Windows path."""

    kind: Literal[CompanionActionKind.FILE_READ] = CompanionActionKind.FILE_READ
    path: str = Field(min_length=3, max_length=1024)
    max_bytes: int = Field(default=1_048_576, ge=1, le=10_485_760)


class FileWriteAction(_ActionModel):
    """Write base64 content below an allowlisted Windows path."""

    kind: Literal[CompanionActionKind.FILE_WRITE] = CompanionActionKind.FILE_WRITE
    path: str = Field(min_length=3, max_length=1024)
    content_base64: str = Field(max_length=13_981_016)
    overwrite: bool = False
    expected_sha256: str | None = Field(default=None, pattern=r"^[A-Fa-f0-9]{64}$")

    @field_validator("content_base64")
    @classmethod
    def validate_base64(cls, value: str) -> str:
        try:
            decoded = base64.b64decode(value, validate=True)
        except (ValueError, binascii.Error) as exc:
            raise ValueError("content_base64 must contain valid base64") from exc
        if len(decoded) > 10_485_760:
            raise ValueError("decoded file content exceeds 10 MiB")
        return value


class OfficeOpenDocumentAction(_ActionModel):
    """Open an allowlisted document in a known Office application."""

    kind: Literal[CompanionActionKind.OFFICE_OPEN_DOCUMENT] = (
        CompanionActionKind.OFFICE_OPEN_DOCUMENT
    )
    application: Literal["word", "powerpoint", "excel"]
    document_path: str = Field(min_length=3, max_length=1024)
    mode: Literal["view", "edit"] = "view"


class OfficeExportPdfAction(_ActionModel):
    """Export an allowlisted Office document to an allowlisted PDF path."""

    kind: Literal[CompanionActionKind.OFFICE_EXPORT_PDF] = (
        CompanionActionKind.OFFICE_EXPORT_PDF
    )
    application: Literal["word", "powerpoint", "excel"]
    source_path: str = Field(min_length=3, max_length=1024)
    output_path: str = Field(min_length=3, max_length=1024)
    overwrite: bool = False


class PowerAutomateDesktopRunAction(_ActionModel):
    """Run one explicitly allowlisted Power Automate Desktop flow."""

    kind: Literal[CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN] = (
        CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN
    )
    flow_name: str = Field(min_length=1, max_length=256)
    inputs: dict[str, Any] = Field(default_factory=dict)
    wait_for_completion: bool = False


class WindowsServiceStatusAction(_ActionModel):
    """Read the status of one explicitly allowlisted Windows service."""

    kind: Literal[CompanionActionKind.WINDOWS_SERVICE_STATUS] = (
        CompanionActionKind.WINDOWS_SERVICE_STATUS
    )
    service_name: str = Field(min_length=1, max_length=256)


class WindowsServiceStartAction(_ActionModel):
    """Start one explicitly allowlisted Windows service."""

    kind: Literal[CompanionActionKind.WINDOWS_SERVICE_START] = (
        CompanionActionKind.WINDOWS_SERVICE_START
    )
    service_name: str = Field(min_length=1, max_length=256)


class WindowsServiceStopAction(_ActionModel):
    """Stop one explicitly allowlisted Windows service."""

    kind: Literal[CompanionActionKind.WINDOWS_SERVICE_STOP] = (
        CompanionActionKind.WINDOWS_SERVICE_STOP
    )
    service_name: str = Field(min_length=1, max_length=256)


class NotificationShowAction(_ActionModel):
    """Show a bounded local Windows notification."""

    kind: Literal[CompanionActionKind.NOTIFICATION_SHOW] = (
        CompanionActionKind.NOTIFICATION_SHOW
    )
    title: str = Field(min_length=1, max_length=128)
    message: str = Field(min_length=1, max_length=2048)


class ClipboardReadTextAction(_ActionModel):
    """Read bounded text from the clipboard with explicit confirmation."""

    kind: Literal[CompanionActionKind.CLIPBOARD_READ_TEXT] = (
        CompanionActionKind.CLIPBOARD_READ_TEXT
    )
    max_characters: int = Field(default=10_000, ge=1, le=100_000)


class ClipboardWriteTextAction(_ActionModel):
    """Write bounded text to the clipboard with explicit confirmation."""

    kind: Literal[CompanionActionKind.CLIPBOARD_WRITE_TEXT] = (
        CompanionActionKind.CLIPBOARD_WRITE_TEXT
    )
    text: str = Field(max_length=100_000)


CompanionAction = Annotated[
    SystemInventoryAction
    | FileListAction
    | FileReadAction
    | FileWriteAction
    | OfficeOpenDocumentAction
    | OfficeExportPdfAction
    | PowerAutomateDesktopRunAction
    | WindowsServiceStatusAction
    | WindowsServiceStartAction
    | WindowsServiceStopAction
    | NotificationShowAction
    | ClipboardReadTextAction
    | ClipboardWriteTextAction,
    Field(discriminator="kind"),
]


class ConfirmationRequirement(StrEnum):
    """How the caller must confirm an action."""

    NONE = "none"
    WHEN_DESTRUCTIVE = "when_destructive"
    ALWAYS = "always"


class ActionPolicy(BaseModel):
    """Policy metadata sent with, and enforced before, an action."""

    model_config = ConfigDict(frozen=True)

    confirmation: ConfirmationRequirement = ConfirmationRequirement.ALWAYS
    destructive: bool = False
    rationale: str = Field(min_length=1, max_length=512)

    @property
    def requires_confirmation(self) -> bool:
        """Return whether this policy requires confirmation."""

        return self.confirmation is ConfirmationRequirement.ALWAYS or (
            self.confirmation is ConfirmationRequirement.WHEN_DESTRUCTIVE
            and self.destructive
        )


def _default_action_policies() -> dict[CompanionActionKind, ActionPolicy]:
    read = ActionPolicy(
        confirmation=ConfirmationRequirement.NONE,
        rationale="Read-only device metadata",
    )
    sensitive_read = ActionPolicy(
        confirmation=ConfirmationRequirement.ALWAYS,
        rationale="May expose user or file content",
    )
    change = ActionPolicy(
        confirmation=ConfirmationRequirement.WHEN_DESTRUCTIVE,
        destructive=True,
        rationale="Changes state on the Windows device",
    )
    return {
        CompanionActionKind.SYSTEM_INVENTORY: read,
        CompanionActionKind.FILE_LIST: sensitive_read,
        CompanionActionKind.FILE_READ: sensitive_read,
        CompanionActionKind.FILE_WRITE: change,
        CompanionActionKind.OFFICE_OPEN_DOCUMENT: change,
        CompanionActionKind.OFFICE_EXPORT_PDF: change,
        CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN: change,
        CompanionActionKind.WINDOWS_SERVICE_STATUS: read,
        CompanionActionKind.WINDOWS_SERVICE_START: change,
        CompanionActionKind.WINDOWS_SERVICE_STOP: change,
        CompanionActionKind.NOTIFICATION_SHOW: change,
        CompanionActionKind.CLIPBOARD_READ_TEXT: sensitive_read,
        CompanionActionKind.CLIPBOARD_WRITE_TEXT: change,
    }


class ConfirmationEvidence(BaseModel):
    """Reference to a confirmation captured by the agent policy layer."""

    model_config = ConfigDict(frozen=True)

    confirmation_id: UUID = Field(default_factory=uuid4)
    action_kind: CompanionActionKind
    confirmed_by: str = Field(min_length=1, max_length=256)
    confirmed_at: datetime
    expires_at: datetime
    purpose: str = Field(min_length=1, max_length=512)
    authorization_reference: str = Field(min_length=1, max_length=512)

    @model_validator(mode="after")
    def validate_window(self) -> ConfirmationEvidence:
        if self.confirmed_at.utcoffset() is None or self.expires_at.utcoffset() is None:
            raise ValueError("confirmation timestamps must include a timezone")
        if self.expires_at <= self.confirmed_at:
            raise ValueError("confirmation must expire after it was granted")
        return self


class DeviceIdentity(BaseModel):
    """Expected Entra and certificate identity of one companion."""

    model_config = ConfigDict(frozen=True)

    device_id: str = Field(pattern=r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
    tenant_id: UUID
    entra_device_id: UUID
    certificate_thumbprint: str = Field(pattern=r"^[A-F0-9]{40}([A-F0-9]{24})?$")

    @field_validator("certificate_thumbprint", mode="before")
    @classmethod
    def normalize_thumbprint(cls, value: Any) -> Any:
        if isinstance(value, str):
            return value.replace(":", "").replace(" ", "").upper()
        return value


class CompanionDevice(BaseModel):
    """Allowlist and expected identity for one Windows laptop."""

    model_config = ConfigDict(frozen=True)

    identity: DeviceIdentity
    display_name: str = Field(min_length=1, max_length=256)
    enabled: bool = True
    allowed_actions: frozenset[CompanionActionKind] = Field(default_factory=frozenset)
    allowed_file_roots: tuple[str, ...] = ()
    allowed_services: frozenset[str] = Field(default_factory=frozenset)
    allowed_desktop_flows: frozenset[str] = Field(default_factory=frozenset)

    @field_validator("allowed_file_roots")
    @classmethod
    def validate_file_roots(cls, values: tuple[str, ...]) -> tuple[str, ...]:
        normalized: list[str] = []
        for value in values:
            path = _normalize_windows_path(value)
            if not ntpath.isabs(path):
                raise ValueError("allowed file roots must be absolute Windows paths")
            normalized.append(path)
        return tuple(normalized)

    @field_validator("allowed_services", "allowed_desktop_flows")
    @classmethod
    def validate_named_allowlist(cls, values: frozenset[str]) -> frozenset[str]:
        if any(not value.strip() or value != value.strip() for value in values):
            raise ValueError("allowlisted names must be non-empty and trimmed")
        return values


class WindowsCompanionSettings(BaseModel):
    """Configuration for the outbound companion control plane."""

    model_config = ConfigDict(frozen=True)

    control_plane_url: HttpUrl
    token_audience: str
    connection_mode: Literal["outbound_relay"] = "outbound_relay"
    timeout_seconds: float = Field(default=30.0, gt=0, le=300)
    tls_profile: str | None = None
    tls_profile_ref: str | None = None
    devices: dict[str, CompanionDevice] = Field(default_factory=dict)
    action_policies: dict[CompanionActionKind, ActionPolicy] = Field(
        default_factory=_default_action_policies
    )

    @field_validator("control_plane_url")
    @classmethod
    def validate_control_plane(cls, value: HttpUrl) -> HttpUrl:
        if value.scheme != "https":
            raise ValueError("companion control plane must use HTTPS")
        if value.username or value.password:
            raise ValueError("control-plane URL cannot contain user credentials")
        if value.path not in {"", "/"} or value.query or value.fragment:
            raise ValueError("control-plane URL must be an origin without a path")
        return value

    @field_validator("token_audience")
    @classmethod
    def validate_audience(cls, value: str) -> str:
        return normalize_audience(value)

    @field_validator("devices")
    @classmethod
    def validate_device_aliases(
        cls, value: dict[str, CompanionDevice]
    ) -> dict[str, CompanionDevice]:
        ids: set[str] = set()
        for alias, device in value.items():
            if not alias or alias != alias.strip() or len(alias) > 128 or "/" in alias:
                raise ValueError("device aliases must be safe, trimmed identifiers")
            if device.identity.device_id in ids:
                raise ValueError("each configured device_id must be unique")
            ids.add(device.identity.device_id)
        return value

    @field_validator("tls_profile")
    @classmethod
    def validate_tls_profile(cls, value: str | None) -> str | None:
        if value is None:
            return None
        selected = value.strip()
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9_.-]{0,127}", selected) is None:
            raise ValueError("companion TLS profile name is invalid")
        return selected

    @field_validator("tls_profile_ref")
    @classmethod
    def validate_tls_profile_ref(cls, value: str | None) -> str | None:
        if value is None:
            return None
        return validate_runtime_secret_reference(value)

    @model_validator(mode="after")
    def require_policies_for_allowed_actions(self) -> WindowsCompanionSettings:
        if self.tls_profile and self.tls_profile_ref:
            raise ValueError("companion TLS selectors are ambiguous")
        missing = {
            action
            for device in self.devices.values()
            for action in device.allowed_actions
            if action not in self.action_policies
        }
        if missing:
            raise ValueError(f"allowed actions are missing policies: {sorted(missing)}")
        return self

    @property
    def api_base_url(self) -> str:
        """Return the stable versioned control-plane API root."""

        return f"{str(self.control_plane_url).rstrip('/')}/v1"


class CompanionActionRequest(BaseModel):
    """Typed action envelope submitted to a Windows companion."""

    model_config = ConfigDict(frozen=True)

    action_id: UUID = Field(default_factory=uuid4)
    action: CompanionAction
    requested_at: datetime = Field(default_factory=lambda: datetime.now(UTC))
    expires_at: datetime = Field(
        default_factory=lambda: datetime.now(UTC) + timedelta(minutes=5)
    )
    idempotency_key: str = Field(
        default_factory=lambda: str(uuid4()), min_length=1, max_length=128
    )
    confirmation: ConfirmationEvidence | None = None

    @field_validator("idempotency_key")
    @classmethod
    def validate_idempotency_key(cls, value: str) -> str:
        if any(char.isspace() or ord(char) < 33 or ord(char) > 126 for char in value):
            raise ValueError("idempotency key must use visible ASCII characters")
        return value

    @model_validator(mode="after")
    def validate_expiry(self) -> CompanionActionRequest:
        if self.requested_at.utcoffset() is None or self.expires_at.utcoffset() is None:
            raise ValueError("action timestamps must include a timezone")
        if self.expires_at <= self.requested_at:
            raise ValueError("action must expire after it was requested")
        return self


class CompanionConnectionStatus(StrEnum):
    """Outbound companion connection state."""

    ONLINE = "online"
    OFFLINE = "offline"
    DEGRADED = "degraded"


class CompanionHealth(BaseModel):
    """Authenticated health reported through the outbound relay."""

    model_config = ConfigDict(frozen=True)

    identity: DeviceIdentity
    status: CompanionConnectionStatus
    authenticated: bool
    outbound_connected: bool
    last_seen_at: datetime
    companion_version: str = Field(min_length=1, max_length=128)
    capabilities: frozenset[CompanionActionKind] = Field(default_factory=frozenset)


class CompanionActionStatus(StrEnum):
    """Action execution lifecycle."""

    ACCEPTED = "accepted"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    REJECTED = "rejected"
    EXPIRED = "expired"
    CANCELED = "canceled"


class CompanionActionReceipt(BaseModel):
    """Receipt returned when an action is accepted or completed quickly."""

    model_config = ConfigDict(frozen=True)

    action_id: UUID
    device_id: str
    status: CompanionActionStatus
    accepted_at: datetime
    status_url: str | None = None


class CompanionActionFailure(BaseModel):
    """Safe structured failure returned by the companion."""

    model_config = ConfigDict(frozen=True)

    code: str = Field(min_length=1, max_length=256)
    message: str = Field(min_length=1, max_length=1000)
    retryable: bool = False


class CompanionActionResult(BaseModel):
    """Final or current result for a companion action."""

    model_config = ConfigDict(frozen=True)

    action_id: UUID
    device_id: str
    status: CompanionActionStatus
    started_at: datetime | None = None
    completed_at: datetime | None = None
    output: dict[str, Any] | None = None
    error: CompanionActionFailure | None = None


class CompanionErrorCode(StrEnum):
    """Stable client-side and upstream error categories."""

    POLICY = "policy_denied"
    AUTHENTICATION = "authentication_failed"
    FORBIDDEN = "forbidden"
    NOT_FOUND = "not_found"
    CONFLICT = "conflict"
    OFFLINE = "device_offline"
    RATE_LIMITED = "rate_limited"
    TIMEOUT = "timeout"
    TRANSPORT = "transport_error"
    INVALID_RESPONSE = "invalid_response"
    UPSTREAM = "upstream_error"


class CompanionError(BaseModel):
    """Safe, normalized control-plane error information."""

    model_config = ConfigDict(frozen=True)

    code: CompanionErrorCode
    message: str
    status_code: int | None = None
    upstream_code: str | None = None
    retry_after_seconds: int | None = None
    correlation_id: str | None = None


class WindowsCompanionClientError(RuntimeError):
    """Raised for local policy, transport, or companion failures."""

    def __init__(self, error: CompanionError):
        self.error = error
        super().__init__(error.message)


class WindowsCompanionClient:
    """Submit typed, allowlisted actions through an authenticated relay."""

    def __init__(
        self,
        settings: WindowsCompanionSettings,
        token_provider: CompanionTokenProvider,
        transport: CompanionHttpTransport,
    ) -> None:
        self.settings = settings
        self._token_provider = token_provider
        self._transport = transport
        # Snapshot validated policy maps so a live client's allowlists cannot be
        # widened by mutating the dictionaries used to build its settings.
        self._devices = dict(settings.devices)
        self._action_policies = dict(settings.action_policies)

    async def get_health(self, device_alias: str) -> CompanionHealth:
        """Read the authenticated outbound connection health for a device."""

        device = self._device(device_alias)
        url = self._device_url(device, "health")
        response = await self._request("GET", url, device=device, expected={200})
        health = self._model_from_response(response, CompanionHealth)
        self._verify_identity(device, health.identity)
        return health

    async def submit_action(
        self,
        device_alias: str,
        request: CompanionActionRequest,
    ) -> CompanionActionReceipt:
        """Validate policy and enqueue one typed action for a companion."""

        device = self._device(device_alias)
        policy = self._authorize_action(device, request)
        payload = {
            "request": request.model_dump(mode="json"),
            "policy": policy.model_dump(mode="json"),
            "expected_device_identity": device.identity.model_dump(mode="json"),
        }
        url = self._device_url(device, "actions")
        response = await self._request(
            "POST",
            url,
            device=device,
            body=self._json_bytes(payload),
            expected={200, 202},
            extra_headers={
                "Idempotency-Key": request.idempotency_key,
                "x-ms-client-request-id": str(request.action_id),
            },
        )
        receipt = self._model_from_response(response, CompanionActionReceipt)
        if receipt.action_id != request.action_id:
            self._raise_invalid_response("Companion returned a different action_id")
        if receipt.device_id != device.identity.device_id:
            self._raise_invalid_response("Companion returned a different device_id")
        return receipt

    async def get_action_result(
        self, device_alias: str, action_id: UUID
    ) -> CompanionActionResult:
        """Read current or final state for a previously submitted action."""

        device = self._device(device_alias)
        url = self._device_url(device, f"actions/{action_id}")
        response = await self._request("GET", url, device=device, expected={200})
        result = self._model_from_response(response, CompanionActionResult)
        if result.action_id != action_id:
            self._raise_invalid_response("Companion returned a different action_id")
        if result.device_id != device.identity.device_id:
            self._raise_invalid_response("Companion returned a different device_id")
        return result

    def _authorize_action(
        self,
        device: CompanionDevice,
        request: CompanionActionRequest,
    ) -> ActionPolicy:
        kind = CompanionActionKind(request.action.kind)
        if kind not in device.allowed_actions:
            self._raise_policy(f"Action {kind.value!r} is not allowed for this device")
        policy = self._action_policies.get(kind)
        if policy is None:
            self._raise_policy(f"Action {kind.value!r} has no configured policy")

        if isinstance(
            request.action,
            (FileListAction, FileReadAction, FileWriteAction),
        ):
            self._require_allowed_path(device, request.action.path)
        elif isinstance(request.action, OfficeOpenDocumentAction):
            self._require_allowed_path(device, request.action.document_path)
        elif isinstance(request.action, OfficeExportPdfAction):
            self._require_allowed_path(device, request.action.source_path)
            self._require_allowed_path(device, request.action.output_path)
        elif isinstance(request.action, PowerAutomateDesktopRunAction):
            if request.action.flow_name.casefold() not in {
                name.casefold() for name in device.allowed_desktop_flows
            }:
                self._raise_policy("Power Automate Desktop flow is not allowlisted")
        elif isinstance(
            request.action,
            (
                WindowsServiceStatusAction,
                WindowsServiceStartAction,
                WindowsServiceStopAction,
            ),
        ) and request.action.service_name.casefold() not in {
            name.casefold() for name in device.allowed_services
        }:
            self._raise_policy("Windows service is not allowlisted")

        if policy.requires_confirmation:
            confirmation = request.confirmation
            if confirmation is None:
                self._raise_policy(f"Action {kind.value!r} requires confirmation")
            now = datetime.now(UTC)
            if confirmation.expires_at <= now:
                self._raise_policy("Action confirmation has expired")
            if confirmation.action_kind is not kind:
                self._raise_policy("Confirmation was issued for a different action")
        if request.expires_at <= datetime.now(UTC):
            self._raise_policy("Action request has expired")
        return policy

    def _require_allowed_path(self, device: CompanionDevice, value: str) -> None:
        candidate = _normalize_windows_path(value)
        if not ntpath.isabs(candidate):
            self._raise_policy("File paths must be absolute Windows paths")
        for root in device.allowed_file_roots:
            try:
                if ntpath.commonpath((candidate, root)) == root:
                    return
            except ValueError:
                continue
        self._raise_policy("File path is outside the device allowlist")

    def _device(self, alias: str) -> CompanionDevice:
        device = self._devices.get(alias)
        if device is None:
            self._raise_policy(f"Device alias {alias!r} is not configured")
        if not device.enabled:
            self._raise_policy(f"Device alias {alias!r} is disabled")
        return device

    def _device_url(self, device: CompanionDevice, suffix: str) -> str:
        device_id = quote(device.identity.device_id, safe="")
        return f"{self.settings.api_base_url}/devices/{device_id}/{suffix}"

    async def _request(
        self,
        method: str,
        url: str,
        *,
        device: CompanionDevice,
        expected: set[int],
        body: bytes | None = None,
        extra_headers: Mapping[str, str] | None = None,
    ) -> CompanionHttpResponse:
        timeout = self.settings.timeout_seconds
        try:
            token = await asyncio.wait_for(
                self._token_provider.get_token(self.settings.token_audience),
                timeout=timeout,
            )
        except TimeoutError as exc:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.TIMEOUT,
                    message=(
                        "Windows companion token acquisition timed out after "
                        f"{timeout:g}s"
                    ),
                )
            ) from exc
        except Exception as exc:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.AUTHENTICATION,
                    message=(
                        "Windows companion token acquisition failed: "
                        f"{type(exc).__name__}"
                    ),
                )
            ) from exc
        if not token or not token.strip():
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.AUTHENTICATION,
                    message="Companion token provider returned no token",
                )
            )

        try:
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}",
                "X-Microsoft-Agent-Device-ID": device.identity.device_id,
                **dict(extra_headers or {}),
            }
            response = await asyncio.wait_for(
                self._transport.request(
                    method,
                    url,
                    headers=headers,
                    body=body,
                    timeout=timeout,
                ),
                timeout=timeout,
            )
        except TimeoutError as exc:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.TIMEOUT,
                    message=f"Windows companion request timed out after {timeout:g}s",
                )
            ) from exc
        except Exception as exc:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.TRANSPORT,
                    message=f"Windows companion transport failed: {type(exc).__name__}",
                )
            ) from exc
        if response.status_code not in expected:
            raise WindowsCompanionClientError(self._error_from_response(response))
        return response

    @staticmethod
    def _verify_identity(expected: CompanionDevice, actual: DeviceIdentity) -> None:
        if actual != expected.identity:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.INVALID_RESPONSE,
                    message="Companion health identity did not match configuration",
                )
            )

    @staticmethod
    def _model_from_response(response: CompanionHttpResponse, model: Any) -> Any:
        try:
            payload = response.json_body()
            return model.model_validate(payload)
        except (TypeError, ValueError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise WindowsCompanionClientError(
                CompanionError(
                    code=CompanionErrorCode.INVALID_RESPONSE,
                    message=f"Invalid companion response for {model.__name__}",
                    status_code=response.status_code,
                )
            ) from exc

    @staticmethod
    def _json_bytes(value: Any) -> bytes:
        try:
            return json.dumps(
                value,
                ensure_ascii=False,
                separators=(",", ":"),
                sort_keys=True,
            ).encode("utf-8")
        except (TypeError, ValueError) as exc:
            raise ValueError("action payload must be JSON serializable") from exc

    @classmethod
    def _error_from_response(cls, response: CompanionHttpResponse) -> CompanionError:
        status = response.status_code
        code = CompanionErrorCode.UPSTREAM
        if status == 401:
            code = CompanionErrorCode.AUTHENTICATION
        elif status == 403:
            code = CompanionErrorCode.FORBIDDEN
        elif status == 404:
            code = CompanionErrorCode.NOT_FOUND
        elif status in {409, 412}:
            code = CompanionErrorCode.CONFLICT
        elif status in {423, 503}:
            code = CompanionErrorCode.OFFLINE
        elif status == 429:
            code = CompanionErrorCode.RATE_LIMITED

        message = f"Windows companion request failed with HTTP {status}"
        upstream_code: str | None = None
        if response.body:
            try:
                payload = response.json_body()
                error = payload.get("error") if isinstance(payload, dict) else None
                if isinstance(error, dict):
                    raw_code = error.get("code")
                    raw_message = error.get("message")
                    if isinstance(raw_code, str):
                        upstream_code = raw_code[:256]
                    if isinstance(raw_message, str) and raw_message.strip():
                        message = raw_message.strip()[:1000]
            except (UnicodeDecodeError, json.JSONDecodeError):
                pass

        retry_after: int | None = None
        raw_retry = cls._header(response.headers, "retry-after")
        if raw_retry and raw_retry.isdigit():
            retry_after = int(raw_retry)
        correlation = cls._header(response.headers, "x-ms-request-id") or cls._header(
            response.headers, "request-id"
        )
        return CompanionError(
            code=code,
            message=message,
            status_code=status,
            upstream_code=upstream_code,
            retry_after_seconds=retry_after,
            correlation_id=correlation,
        )

    @staticmethod
    def _header(headers: Mapping[str, str], name: str) -> str | None:
        wanted = name.casefold()
        for key, value in headers.items():
            if key.casefold() == wanted:
                return value
        return None

    @staticmethod
    def _raise_policy(message: str) -> Never:
        raise WindowsCompanionClientError(
            CompanionError(code=CompanionErrorCode.POLICY, message=message)
        )

    @staticmethod
    def _raise_invalid_response(message: str) -> Never:
        raise WindowsCompanionClientError(
            CompanionError(code=CompanionErrorCode.INVALID_RESPONSE, message=message)
        )


def _normalize_windows_path(value: str) -> str:
    if not value or "\x00" in value:
        raise ValueError("Windows path must be non-empty and contain no NUL")
    return ntpath.normcase(ntpath.normpath(value.strip()))


__all__ = [
    "ActionPolicy",
    "ClipboardReadTextAction",
    "ClipboardWriteTextAction",
    "CompanionAction",
    "CompanionActionFailure",
    "CompanionActionKind",
    "CompanionActionReceipt",
    "CompanionActionRequest",
    "CompanionActionResult",
    "CompanionActionStatus",
    "CompanionConnectionStatus",
    "CompanionDevice",
    "CompanionError",
    "CompanionErrorCode",
    "CompanionHealth",
    "CompanionHttpResponse",
    "CompanionHttpTransport",
    "CompanionTokenProvider",
    "ConfirmationEvidence",
    "ConfirmationRequirement",
    "DeviceIdentity",
    "FileListAction",
    "FileReadAction",
    "FileWriteAction",
    "HttpxCompanionTransport",
    "NotificationShowAction",
    "OfficeExportPdfAction",
    "OfficeOpenDocumentAction",
    "PowerAutomateDesktopRunAction",
    "SystemInventoryAction",
    "WindowsCompanionClient",
    "WindowsCompanionClientError",
    "WindowsCompanionSettings",
    "WindowsServiceStartAction",
    "WindowsServiceStatusAction",
    "WindowsServiceStopAction",
]
