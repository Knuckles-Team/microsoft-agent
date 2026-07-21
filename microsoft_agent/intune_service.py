"""Policy-enforced Microsoft Intune managed-device client.

This module intentionally exposes only fixed Microsoft Graph v1.0 endpoints.  It
does not accept arbitrary URLs, paths, actions, headers, or request bodies.
Authentication and HTTP transport are injected so production code can reuse the
package's credential stack while tests remain network-free.
"""

from __future__ import annotations

import asyncio
import json
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from enum import StrEnum
from typing import Any, Literal, Protocol, cast
from uuid import UUID, uuid4

from pydantic import (
    AwareDatetime,
    BaseModel,
    ConfigDict,
    Field,
    SecretStr,
    field_validator,
    model_validator,
)

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_AUDIENCE = "https://graph.microsoft.com"


class DeviceAction(StrEnum):
    """Supported or explicitly classified Intune managed-device actions."""

    SYNC_DEVICE = "syncDevice"
    REBOOT_NOW = "rebootNow"
    REMOTE_LOCK = "remoteLock"
    SHUT_DOWN = "shutDown"
    WINDOWS_DEFENDER_SCAN = "windowsDefenderScan"
    ROTATE_BITLOCKER_KEYS = "rotateBitLockerKeys"


class IntuneErrorCode(StrEnum):
    """Stable machine-readable service error codes."""

    AUTHENTICATION_FAILED = "authentication_failed"
    CONFIRMATION_INVALID = "confirmation_invalid"
    DEVICE_NOT_ALLOWED = "device_not_allowed"
    ACTION_NOT_ALLOWED = "action_not_allowed"
    CAPABILITY_UNSUPPORTED = "capability_unsupported"
    DETECTED_APPS_NOT_ALLOWED = "detected_apps_not_allowed"
    GRAPH_REQUEST_FAILED = "graph_request_failed"
    IDEMPOTENCY_CONFLICT = "idempotency_conflict"
    INVALID_DEVICE_ID = "invalid_device_id"
    INVALID_RESPONSE = "invalid_response"


class GraphAccessToken(BaseModel):
    """Access token whose provider attests the Microsoft Graph audience."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    access_token: SecretStr
    audience: Literal["https://graph.microsoft.com"] = "https://graph.microsoft.com"

    @field_validator("access_token")
    @classmethod
    def validate_access_token(cls, value: SecretStr) -> SecretStr:
        token = value.get_secret_value().strip()
        if not token:
            raise ValueError("access_token must not be empty")
        if token.lower().startswith("bearer "):
            raise ValueError("access_token must not include the Bearer scheme")
        return SecretStr(token)


class GraphTokenProvider(Protocol):
    """Async provider for audience-bound Microsoft Graph tokens."""

    async def get_token(self, audience: str) -> GraphAccessToken:
        """Return an access token for exactly ``audience``."""


class HttpResponse(Protocol):
    """Minimal response contract required from an injected HTTP client."""

    status_code: int
    headers: Mapping[str, str]

    def json(self) -> Any:
        """Decode the response JSON body."""


class AsyncHttpClient(Protocol):
    """Minimal async request contract compatible with common HTTP clients."""

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, str] | None = None,
        json: Any = None,
    ) -> HttpResponse:
        """Send one HTTP request without following redirects."""


class IntuneServiceSettings(BaseModel):
    """Fail-closed policy and endpoint configuration for :class:`IntuneService`."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    graph_base_url: str = GRAPH_BASE_URL
    graph_audience: str = GRAPH_AUDIENCE
    allowed_device_ids: frozenset[UUID] = Field(min_length=1)
    allowed_actions: frozenset[DeviceAction] = Field(default_factory=frozenset)
    allow_tenant_detected_apps: bool = False
    max_confirmation_lifetime_seconds: int = Field(default=600, ge=30, le=3600)
    clock_skew_seconds: int = Field(default=30, ge=0, le=300)

    @field_validator("graph_base_url")
    @classmethod
    def require_v1_graph_url(cls, value: str) -> str:
        if value.rstrip("/") != GRAPH_BASE_URL:
            raise ValueError(
                "graph_base_url must be the HTTPS Microsoft Graph v1.0 endpoint"
            )
        return GRAPH_BASE_URL

    @field_validator("graph_audience")
    @classmethod
    def require_graph_audience(cls, value: str) -> str:
        if value.rstrip("/") != GRAPH_AUDIENCE:
            raise ValueError("graph_audience must be Microsoft Graph")
        return GRAPH_AUDIENCE


class ConfirmationEvidence(BaseModel):
    """Auditable approval bound to one device, action, and idempotency key."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    confirmation_id: UUID
    device_id: UUID
    action: DeviceAction
    approved: Literal[True]
    confirmed_by: str = Field(min_length=1, max_length=256)
    reason: str = Field(min_length=1, max_length=1024)
    confirmed_at: AwareDatetime
    expires_at: AwareDatetime
    idempotency_key: UUID
    correlation_id: UUID
    destructive_action_acknowledged: bool = False

    @model_validator(mode="after")
    def validate_time_order(self) -> ConfirmationEvidence:
        if self.expires_at <= self.confirmed_at:
            raise ValueError("expires_at must be later than confirmed_at")
        return self


class ManagedDevice(BaseModel):
    """Selected managed-device fields with forward-compatible extra properties."""

    model_config = ConfigDict(populate_by_name=True, extra="allow", frozen=True)

    id: UUID
    device_name: str | None = Field(default=None, alias="deviceName")
    operating_system: str | None = Field(default=None, alias="operatingSystem")
    os_version: str | None = Field(default=None, alias="osVersion")
    compliance_state: str | None = Field(default=None, alias="complianceState")
    last_sync_date_time: datetime | None = Field(default=None, alias="lastSyncDateTime")
    azure_ad_device_id: str | None = Field(default=None, alias="azureADDeviceId")
    serial_number: str | None = Field(default=None, alias="serialNumber")
    manufacturer: str | None = None
    model: str | None = None
    is_encrypted: bool | None = Field(default=None, alias="isEncrypted")
    owner_type: str | None = Field(default=None, alias="managedDeviceOwnerType")
    user_principal_name: str | None = Field(default=None, alias="userPrincipalName")


class ManagedDevicePage(BaseModel):
    """Allowlist-filtered page of Intune managed devices."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    value: tuple[ManagedDevice, ...]
    truncated: bool = False


class DetectedApp(BaseModel):
    """Application discovered by Intune inventory."""

    model_config = ConfigDict(populate_by_name=True, extra="allow", frozen=True)

    id: UUID
    display_name: str | None = Field(default=None, alias="displayName")
    version: str | None = None
    size_in_bytes: int | None = Field(default=None, alias="sizeInByte", ge=0)
    device_count: int | None = Field(default=None, alias="deviceCount", ge=0)
    publisher: str | None = None
    platform: str | None = None


class DetectedAppPage(BaseModel):
    """Page of tenant-wide detected applications."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    value: tuple[DetectedApp, ...]
    truncated: bool = False


class ActionCapability(BaseModel):
    """Static security and support classification for a device action."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    action: DeviceAction
    endpoint_suffix: str
    supported: bool
    api_version: Literal["v1.0"] | None
    mutating: Literal[True] = True
    confirmation_required: Literal[True] = True
    destructive: bool
    disruptive: bool
    graph_idempotency_documented: Literal[False] = False
    unsupported_reason: str | None = None


class DeviceActionResult(BaseModel):
    """Accepted Intune action with audit and client-side deduplication metadata."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    status: Literal["accepted"] = "accepted"
    device_id: UUID
    action: DeviceAction
    destructive: bool
    disruptive: bool
    correlation_id: UUID
    idempotency_key: UUID
    graph_request_id: str | None = None
    accepted_at: AwareDatetime
    replayed: bool = False


class IntuneError(BaseModel):
    """Structured service error suitable for API and audit boundaries."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    code: IntuneErrorCode
    message: str
    http_status: int | None = None
    graph_code: str | None = None
    correlation_id: UUID | None = None
    graph_request_id: str | None = None
    retryable: bool = False


class IntuneServiceError(RuntimeError):
    """Base exception carrying a typed, stable error payload."""

    def __init__(self, error: IntuneError) -> None:
        self.error = error
        super().__init__(error.message)


class AuthenticationError(IntuneServiceError):
    """Token provider returned unusable or incorrectly scoped credentials."""


class PolicyViolationError(IntuneServiceError):
    """The device, action, or inventory capability is not explicitly allowed."""


class ConfirmationError(IntuneServiceError):
    """Mutation confirmation evidence is missing, stale, or mismatched."""


class UnsupportedCapabilityError(IntuneServiceError):
    """The requested action is deliberately unavailable in Graph v1.0."""


class GraphRequestError(IntuneServiceError):
    """Microsoft Graph returned an unsuccessful response."""


class ResponseValidationError(IntuneServiceError):
    """Microsoft Graph returned a response with an unexpected shape."""


class IdempotencyConflictError(IntuneServiceError):
    """An idempotency key was reused for a different operation."""


_CAPABILITIES: dict[DeviceAction, ActionCapability] = {
    DeviceAction.SYNC_DEVICE: ActionCapability(
        action=DeviceAction.SYNC_DEVICE,
        endpoint_suffix="syncDevice",
        supported=True,
        api_version="v1.0",
        destructive=False,
        disruptive=False,
    ),
    DeviceAction.REBOOT_NOW: ActionCapability(
        action=DeviceAction.REBOOT_NOW,
        endpoint_suffix="rebootNow",
        supported=True,
        api_version="v1.0",
        destructive=True,
        disruptive=True,
    ),
    DeviceAction.REMOTE_LOCK: ActionCapability(
        action=DeviceAction.REMOTE_LOCK,
        endpoint_suffix="remoteLock",
        supported=True,
        api_version="v1.0",
        destructive=False,
        disruptive=True,
    ),
    DeviceAction.SHUT_DOWN: ActionCapability(
        action=DeviceAction.SHUT_DOWN,
        endpoint_suffix="shutDown",
        supported=True,
        api_version="v1.0",
        destructive=True,
        disruptive=True,
    ),
    DeviceAction.WINDOWS_DEFENDER_SCAN: ActionCapability(
        action=DeviceAction.WINDOWS_DEFENDER_SCAN,
        endpoint_suffix="windowsDefenderScan",
        supported=True,
        api_version="v1.0",
        destructive=False,
        disruptive=False,
    ),
    DeviceAction.ROTATE_BITLOCKER_KEYS: ActionCapability(
        action=DeviceAction.ROTATE_BITLOCKER_KEYS,
        endpoint_suffix="rotateBitLockerKeys",
        supported=False,
        api_version=None,
        destructive=True,
        disruptive=False,
        unsupported_reason=(
            "rotateBitLockerKeys is not documented for Microsoft Graph v1.0; "
            "the beta endpoint is intentionally disabled"
        ),
    ),
}

_MANAGED_DEVICE_SELECT = ",".join(
    (
        "id",
        "deviceName",
        "operatingSystem",
        "osVersion",
        "complianceState",
        "lastSyncDateTime",
        "azureADDeviceId",
        "serialNumber",
        "manufacturer",
        "model",
        "isEncrypted",
        "managedDeviceOwnerType",
        "userPrincipalName",
    )
)
_DETECTED_APP_SELECT = ",".join(
    (
        "id",
        "displayName",
        "version",
        "sizeInByte",
        "deviceCount",
        "publisher",
        "platform",
    )
)
_NO_BODY = object()


@dataclass(frozen=True)
class _IdempotencyRecord:
    fingerprint: str
    result: DeviceActionResult


class IntuneService:
    """Async, allowlist-enforced client for selected Intune v1.0 operations."""

    def __init__(
        self,
        settings: IntuneServiceSettings,
        http_client: AsyncHttpClient,
        token_provider: GraphTokenProvider,
        *,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.settings = settings
        self._http_client = http_client
        self._token_provider = token_provider
        self._clock = clock or (lambda: datetime.now(UTC))
        self._idempotency_records: dict[UUID, _IdempotencyRecord] = {}
        self._action_lock = asyncio.Lock()

    @staticmethod
    def capabilities() -> tuple[ActionCapability, ...]:
        """Return immutable action classifications in stable enum order."""
        return tuple(_CAPABILITIES[action] for action in DeviceAction)

    async def list_managed_devices(self) -> ManagedDevicePage:
        """List only explicitly allowlisted managed devices."""
        allowed_ids = sorted(self.settings.allowed_device_ids, key=str)
        filter_value = " or ".join(f"id eq '{device_id}'" for device_id in allowed_ids)
        correlation_id = uuid4()
        payload, _ = await self._request_json(
            "GET",
            "/deviceManagement/managedDevices",
            expected_status=200,
            correlation_id=correlation_id,
            params={
                "$filter": filter_value,
                "$select": _MANAGED_DEVICE_SELECT,
                "$top": str(len(allowed_ids)),
            },
        )
        values = self._collection_values(payload, correlation_id)
        devices = tuple(
            device
            for item in values
            if (device := self._validate_device(item, correlation_id)).id
            in self.settings.allowed_device_ids
        )
        return ManagedDevicePage(
            value=devices, truncated=bool(payload.get("@odata.nextLink"))
        )

    async def get_managed_device(self, device_id: UUID | str) -> ManagedDevice:
        """Get one explicitly allowlisted managed device."""
        validated_id = self._require_allowed_device(device_id)
        correlation_id = uuid4()
        payload, _ = await self._request_json(
            "GET",
            f"/deviceManagement/managedDevices/{validated_id}",
            expected_status=200,
            correlation_id=correlation_id,
            params={"$select": _MANAGED_DEVICE_SELECT},
        )
        item = (
            payload.get("value") if isinstance(payload.get("value"), dict) else payload
        )
        device = self._validate_device(item, correlation_id)
        if device.id != validated_id:
            raise self._response_error(
                "Graph returned a managed device with an unexpected ID",
                correlation_id,
            )
        return device

    async def list_detected_apps(self) -> DetectedAppPage:
        """List tenant-wide detected apps when that inventory scope is opted in."""
        if not self.settings.allow_tenant_detected_apps:
            raise PolicyViolationError(
                IntuneError(
                    code=IntuneErrorCode.DETECTED_APPS_NOT_ALLOWED,
                    message="Tenant-wide detected-app inventory is not allowed",
                )
            )
        correlation_id = uuid4()
        payload, _ = await self._request_json(
            "GET",
            "/deviceManagement/detectedApps",
            expected_status=200,
            correlation_id=correlation_id,
            params={"$select": _DETECTED_APP_SELECT},
        )
        values = self._collection_values(payload, correlation_id)
        try:
            apps = tuple(DetectedApp.model_validate(item) for item in values)
        except Exception as exc:
            raise self._response_error(
                "Graph returned an invalid detected-app object", correlation_id
            ) from exc
        return DetectedAppPage(
            value=apps, truncated=bool(payload.get("@odata.nextLink"))
        )

    async def sync_device(
        self,
        device_id: UUID | str,
        *,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Request an Intune device sync."""
        return await self._execute_action(
            device_id, DeviceAction.SYNC_DEVICE, confirmation=confirmation
        )

    async def reboot_now(
        self,
        device_id: UUID | str,
        *,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Immediately request a device reboot."""
        return await self._execute_action(
            device_id, DeviceAction.REBOOT_NOW, confirmation=confirmation
        )

    async def remote_lock(
        self,
        device_id: UUID | str,
        *,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Request a remote device lock."""
        return await self._execute_action(
            device_id, DeviceAction.REMOTE_LOCK, confirmation=confirmation
        )

    async def shut_down(
        self,
        device_id: UUID | str,
        *,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Immediately request a device shutdown."""
        return await self._execute_action(
            device_id, DeviceAction.SHUT_DOWN, confirmation=confirmation
        )

    async def windows_defender_scan(
        self,
        device_id: UUID | str,
        *,
        quick_scan: bool,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Request a quick or full Microsoft Defender scan."""
        return await self._execute_action(
            device_id,
            DeviceAction.WINDOWS_DEFENDER_SCAN,
            confirmation=confirmation,
            body={"quickScan": quick_scan},
        )

    async def rotate_bitlocker_keys(
        self,
        device_id: UUID | str,
        *,
        confirmation: ConfirmationEvidence | None,
    ) -> DeviceActionResult:
        """Fail closed because key rotation is not documented in Graph v1.0."""
        return await self._execute_action(
            device_id,
            DeviceAction.ROTATE_BITLOCKER_KEYS,
            confirmation=confirmation,
        )

    async def _execute_action(
        self,
        device_id: UUID | str,
        action: DeviceAction,
        *,
        confirmation: ConfirmationEvidence | None,
        body: Any = _NO_BODY,
    ) -> DeviceActionResult:
        validated_id = self._require_allowed_device(device_id)
        self._require_allowed_action(action)
        capability = _CAPABILITIES[action]
        if not capability.supported:
            raise UnsupportedCapabilityError(
                IntuneError(
                    code=IntuneErrorCode.CAPABILITY_UNSUPPORTED,
                    message=cast(str, capability.unsupported_reason),
                )
            )
        evidence = self._validate_confirmation(confirmation, validated_id, capability)
        fingerprint = self._action_fingerprint(validated_id, action, body)

        async with self._action_lock:
            previous = self._idempotency_records.get(evidence.idempotency_key)
            if previous:
                if previous.fingerprint != fingerprint:
                    raise IdempotencyConflictError(
                        IntuneError(
                            code=IntuneErrorCode.IDEMPOTENCY_CONFLICT,
                            message=(
                                "The idempotency key was already used for a different "
                                "device action"
                            ),
                            correlation_id=evidence.correlation_id,
                        )
                    )
                return previous.result.model_copy(update={"replayed": True})

            _, response = await self._request_json(
                "POST",
                (
                    f"/deviceManagement/managedDevices/{validated_id}/"
                    f"{capability.endpoint_suffix}"
                ),
                expected_status=204,
                correlation_id=evidence.correlation_id,
                idempotency_key=evidence.idempotency_key,
                body=body,
            )
            result = DeviceActionResult(
                device_id=validated_id,
                action=action,
                destructive=capability.destructive,
                disruptive=capability.disruptive,
                correlation_id=evidence.correlation_id,
                idempotency_key=evidence.idempotency_key,
                graph_request_id=self._header(response.headers, "request-id"),
                accepted_at=self._now(),
            )
            self._idempotency_records[evidence.idempotency_key] = _IdempotencyRecord(
                fingerprint=fingerprint, result=result
            )
            return result

    def _require_allowed_device(self, device_id: UUID | str) -> UUID:
        try:
            validated_id = device_id if isinstance(device_id, UUID) else UUID(device_id)
        except (TypeError, ValueError, AttributeError) as exc:
            raise PolicyViolationError(
                IntuneError(
                    code=IntuneErrorCode.INVALID_DEVICE_ID,
                    message="managed device ID must be a UUID",
                )
            ) from exc
        if validated_id not in self.settings.allowed_device_ids:
            raise PolicyViolationError(
                IntuneError(
                    code=IntuneErrorCode.DEVICE_NOT_ALLOWED,
                    message="The managed device is not in the explicit allowlist",
                )
            )
        return validated_id

    def _require_allowed_action(self, action: DeviceAction) -> None:
        if action not in self.settings.allowed_actions:
            raise PolicyViolationError(
                IntuneError(
                    code=IntuneErrorCode.ACTION_NOT_ALLOWED,
                    message=f"The {action.value} action is not in the explicit allowlist",
                )
            )

    def _validate_confirmation(
        self,
        confirmation: ConfirmationEvidence | None,
        device_id: UUID,
        capability: ActionCapability,
    ) -> ConfirmationEvidence:
        message: str | None = None
        if confirmation is None:
            message = "Confirmation evidence is required for every device action"
        elif confirmation.device_id != device_id:
            message = "Confirmation evidence is for a different device"
        elif confirmation.action != capability.action:
            message = "Confirmation evidence is for a different action"
        else:
            now = self._now()
            skew = timedelta(seconds=self.settings.clock_skew_seconds)
            lifetime = confirmation.expires_at - confirmation.confirmed_at
            if confirmation.confirmed_at > now + skew:
                message = "Confirmation evidence is dated in the future"
            elif now > confirmation.expires_at + skew:
                message = "Confirmation evidence has expired"
            elif lifetime > timedelta(
                seconds=self.settings.max_confirmation_lifetime_seconds
            ):
                message = "Confirmation evidence lifetime exceeds policy"
            elif (
                capability.destructive
                and not confirmation.destructive_action_acknowledged
            ):
                message = "The destructive action was not explicitly acknowledged"
        if message:
            raise ConfirmationError(
                IntuneError(
                    code=IntuneErrorCode.CONFIRMATION_INVALID,
                    message=message,
                    correlation_id=(
                        confirmation.correlation_id if confirmation else None
                    ),
                )
            )
        return cast(ConfirmationEvidence, confirmation)

    async def _request_json(
        self,
        method: Literal["GET", "POST"],
        path: str,
        *,
        expected_status: int,
        correlation_id: UUID,
        params: Mapping[str, str] | None = None,
        idempotency_key: UUID | None = None,
        body: Any = _NO_BODY,
    ) -> tuple[dict[str, Any], HttpResponse]:
        headers = await self._headers(correlation_id, idempotency_key, body)
        url = f"{self.settings.graph_base_url}{path}"
        request_kwargs: dict[str, Any] = {"headers": headers}
        if params:
            request_kwargs["params"] = params
        if body is not _NO_BODY:
            request_kwargs["json"] = body
        try:
            response = await self._http_client.request(method, url, **request_kwargs)
        except IntuneServiceError:
            raise
        except Exception as exc:
            raise GraphRequestError(
                IntuneError(
                    code=IntuneErrorCode.GRAPH_REQUEST_FAILED,
                    message="Microsoft Graph request failed before a response was received",
                    correlation_id=correlation_id,
                    retryable=True,
                )
            ) from exc

        if response.status_code != expected_status:
            raise self._graph_error(response, correlation_id)
        if expected_status == 204:
            return {}, response
        try:
            payload = response.json()
        except Exception as exc:
            raise self._response_error(
                "Microsoft Graph returned invalid JSON", correlation_id
            ) from exc
        if not isinstance(payload, dict):
            raise self._response_error(
                "Microsoft Graph returned a non-object JSON response", correlation_id
            )
        return payload, response

    async def _headers(
        self,
        correlation_id: UUID,
        idempotency_key: UUID | None,
        body: Any,
    ) -> dict[str, str]:
        try:
            supplied = await self._token_provider.get_token(
                self.settings.graph_audience
            )
            token = GraphAccessToken.model_validate(supplied)
        except Exception as exc:
            raise AuthenticationError(
                IntuneError(
                    code=IntuneErrorCode.AUTHENTICATION_FAILED,
                    message="Unable to acquire an audience-bound Microsoft Graph token",
                    correlation_id=correlation_id,
                )
            ) from exc
        if token.audience != self.settings.graph_audience:
            raise AuthenticationError(
                IntuneError(
                    code=IntuneErrorCode.AUTHENTICATION_FAILED,
                    message="Token audience is not Microsoft Graph",
                    correlation_id=correlation_id,
                )
            )
        headers = {
            "Authorization": f"Bearer {token.access_token.get_secret_value()}",
            "Accept": "application/json",
            "client-request-id": str(correlation_id),
            "return-client-request-id": "true",
        }
        if idempotency_key:
            headers["Idempotency-Key"] = str(idempotency_key)
        if body is not _NO_BODY:
            headers["Content-Type"] = "application/json"
        return headers

    def _graph_error(
        self, response: HttpResponse, correlation_id: UUID
    ) -> GraphRequestError:
        graph_code: str | None = None
        message = "Microsoft Graph rejected the request"
        graph_request_id = self._header(response.headers, "request-id")
        try:
            payload = response.json()
            graph_error = payload.get("error") if isinstance(payload, dict) else None
            if isinstance(graph_error, dict):
                if isinstance(graph_error.get("code"), str):
                    graph_code = graph_error["code"]
                if isinstance(graph_error.get("message"), str):
                    message = graph_error["message"]
                inner_error = graph_error.get("innerError") or graph_error.get(
                    "innererror"
                )
                if isinstance(inner_error, dict):
                    inner_request_id = inner_error.get("request-id")
                    if isinstance(inner_request_id, str):
                        graph_request_id = inner_request_id
        except Exception:
            pass
        return GraphRequestError(
            IntuneError(
                code=IntuneErrorCode.GRAPH_REQUEST_FAILED,
                message=message,
                http_status=response.status_code,
                graph_code=graph_code,
                correlation_id=correlation_id,
                graph_request_id=graph_request_id,
                retryable=response.status_code in {408, 429, 500, 502, 503, 504},
            )
        )

    @staticmethod
    def _collection_values(payload: dict[str, Any], correlation_id: UUID) -> list[Any]:
        values = payload.get("value")
        if not isinstance(values, list):
            raise ResponseValidationError(
                IntuneError(
                    code=IntuneErrorCode.INVALID_RESPONSE,
                    message="Microsoft Graph collection response has no value array",
                    correlation_id=correlation_id,
                )
            )
        return values

    @staticmethod
    def _validate_device(item: Any, correlation_id: UUID) -> ManagedDevice:
        try:
            return ManagedDevice.model_validate(item)
        except Exception as exc:
            raise ResponseValidationError(
                IntuneError(
                    code=IntuneErrorCode.INVALID_RESPONSE,
                    message="Microsoft Graph returned an invalid managed-device object",
                    correlation_id=correlation_id,
                )
            ) from exc

    @staticmethod
    def _response_error(message: str, correlation_id: UUID) -> ResponseValidationError:
        return ResponseValidationError(
            IntuneError(
                code=IntuneErrorCode.INVALID_RESPONSE,
                message=message,
                correlation_id=correlation_id,
            )
        )

    @staticmethod
    def _action_fingerprint(device_id: UUID, action: DeviceAction, body: Any) -> str:
        normalized_body = None if body is _NO_BODY else body
        return json.dumps(
            {
                "device_id": str(device_id),
                "action": action.value,
                "body": normalized_body,
            },
            sort_keys=True,
            separators=(",", ":"),
        )

    @staticmethod
    def _header(headers: Mapping[str, str], name: str) -> str | None:
        expected = name.lower()
        for key, value in headers.items():
            if key.lower() == expected:
                return value
        return None

    def _now(self) -> datetime:
        now = self._clock()
        if now.tzinfo is None or now.utcoffset() is None:
            raise RuntimeError(
                "IntuneService clock must return a timezone-aware datetime"
            )
        return now


__all__ = [
    "ActionCapability",
    "AsyncHttpClient",
    "AuthenticationError",
    "ConfirmationError",
    "ConfirmationEvidence",
    "DetectedApp",
    "DetectedAppPage",
    "DeviceAction",
    "DeviceActionResult",
    "GraphAccessToken",
    "GraphRequestError",
    "GraphTokenProvider",
    "IdempotencyConflictError",
    "IntuneError",
    "IntuneErrorCode",
    "IntuneService",
    "IntuneServiceError",
    "IntuneServiceSettings",
    "ManagedDevice",
    "ManagedDevicePage",
    "PolicyViolationError",
    "ResponseValidationError",
    "UnsupportedCapabilityError",
]
