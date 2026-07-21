"""Network-free tests for the policy-enforced Intune service."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import Any, cast
from uuid import UUID

import pytest
from pydantic import ValidationError

from microsoft_agent.intune_service import (
    AsyncHttpClient,
    AuthenticationError,
    ConfirmationError,
    ConfirmationEvidence,
    DeviceAction,
    GraphAccessToken,
    GraphRequestError,
    IdempotencyConflictError,
    IntuneErrorCode,
    IntuneService,
    IntuneServiceSettings,
    PolicyViolationError,
    UnsupportedCapabilityError,
)

DEVICE_ID = UUID("11111111-1111-4111-8111-111111111111")
SECOND_DEVICE_ID = UUID("22222222-2222-4222-8222-222222222222")
UNAUTHORIZED_DEVICE_ID = UUID("33333333-3333-4333-8333-333333333333")
APP_ID = UUID("aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa")
NOW = datetime(2026, 7, 14, 15, 0, tzinfo=UTC)


class FakeResponse:
    def __init__(
        self,
        status_code: int,
        payload: Any = None,
        headers: dict[str, str] | None = None,
        *,
        json_error: Exception | None = None,
    ) -> None:
        self.status_code = status_code
        self.payload = payload
        self.headers = headers or {}
        self.json_error = json_error

    def json(self) -> Any:
        if self.json_error:
            raise self.json_error
        return self.payload


class FakeHttpClient:
    def __init__(self, *responses: FakeResponse) -> None:
        self.responses = list(responses)
        self.calls: list[tuple[str, str, dict[str, Any]]] = []

    async def request(self, method: str, url: str, **kwargs: Any) -> FakeResponse:
        self.calls.append((method, url, kwargs))
        if not self.responses:
            raise AssertionError("Unexpected HTTP request")
        return self.responses.pop(0)


class FakeTokenProvider:
    def __init__(self, token: GraphAccessToken | dict[str, Any] | None = None) -> None:
        self.token = token or GraphAccessToken(access_token="secret-token")
        self.audiences: list[str] = []

    async def get_token(self, audience: str) -> Any:
        self.audiences.append(audience)
        return self.token


def settings(
    *,
    allowed_actions: frozenset[DeviceAction] | None = None,
    allow_tenant_detected_apps: bool = False,
) -> IntuneServiceSettings:
    return IntuneServiceSettings(
        allowed_device_ids=frozenset({DEVICE_ID, SECOND_DEVICE_ID}),
        allowed_actions=(
            allowed_actions if allowed_actions is not None else frozenset(DeviceAction)
        ),
        allow_tenant_detected_apps=allow_tenant_detected_apps,
    )


def confirmation(
    action: DeviceAction,
    *,
    device_id: UUID = DEVICE_ID,
    idempotency_key: UUID = UUID("44444444-4444-4444-8444-444444444444"),
    correlation_id: UUID = UUID("55555555-5555-4555-8555-555555555555"),
    confirmed_at: datetime = NOW - timedelta(minutes=1),
    expires_at: datetime = NOW + timedelta(minutes=4),
    destructive_acknowledged: bool = False,
) -> ConfirmationEvidence:
    return ConfirmationEvidence(
        confirmation_id=UUID("66666666-6666-4666-8666-666666666666"),
        device_id=device_id,
        action=action,
        approved=True,
        confirmed_by="admin@example.com",
        reason="Approved device maintenance",
        confirmed_at=confirmed_at,
        expires_at=expires_at,
        idempotency_key=idempotency_key,
        correlation_id=correlation_id,
        destructive_action_acknowledged=destructive_acknowledged,
    )


def service(
    http: FakeHttpClient,
    *,
    service_settings: IntuneServiceSettings | None = None,
    tokens: FakeTokenProvider | None = None,
) -> IntuneService:
    return IntuneService(
        service_settings or settings(),
        cast(AsyncHttpClient, http),
        tokens or FakeTokenProvider(),
        clock=lambda: NOW,
    )


def test_settings_are_https_v1_graph_only_and_require_devices() -> None:
    with pytest.raises(ValidationError):
        IntuneServiceSettings(
            graph_base_url="http://graph.microsoft.com/v1.0",
            allowed_device_ids=frozenset({DEVICE_ID}),
        )
    with pytest.raises(ValidationError):
        IntuneServiceSettings(
            graph_base_url="https://example.com/v1.0",
            allowed_device_ids=frozenset({DEVICE_ID}),
        )
    with pytest.raises(ValidationError):
        IntuneServiceSettings(
            graph_audience="https://management.azure.com",
            allowed_device_ids=frozenset({DEVICE_ID}),
        )
    with pytest.raises(ValidationError):
        IntuneServiceSettings(allowed_device_ids=frozenset())


@pytest.mark.asyncio
async def test_list_devices_uses_fixed_endpoint_and_filters_to_allowlist() -> None:
    http = FakeHttpClient(
        FakeResponse(
            200,
            {
                "value": [
                    {
                        "id": str(DEVICE_ID),
                        "deviceName": "laptop-one",
                        "operatingSystem": "Windows",
                    },
                    {
                        "id": str(UNAUTHORIZED_DEVICE_ID),
                        "deviceName": "not-allowed",
                    },
                ],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/opaque",
            },
        )
    )
    tokens = FakeTokenProvider()
    client = service(http, tokens=tokens)

    result = await client.list_managed_devices()

    assert [device.id for device in result.value] == [DEVICE_ID]
    assert result.truncated is True
    method, url, kwargs = http.calls[0]
    assert method == "GET"
    assert url == "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
    assert kwargs["params"]["$filter"] == (
        f"id eq '{DEVICE_ID}' or id eq '{SECOND_DEVICE_ID}'"
    )
    assert kwargs["params"]["$top"] == "2"
    assert kwargs["headers"]["Authorization"] == "Bearer secret-token"
    assert tokens.audiences == ["https://graph.microsoft.com"]


@pytest.mark.asyncio
async def test_get_device_is_allowlisted_and_id_bound() -> None:
    http = FakeHttpClient(
        FakeResponse(200, {"id": str(DEVICE_ID), "deviceName": "laptop-one"})
    )
    client = service(http)

    result = await client.get_managed_device(DEVICE_ID)

    assert result.id == DEVICE_ID
    assert http.calls[0][1].endswith(f"/deviceManagement/managedDevices/{DEVICE_ID}")

    with pytest.raises(PolicyViolationError) as exc_info:
        await client.get_managed_device(UNAUTHORIZED_DEVICE_ID)
    assert exc_info.value.error.code == IntuneErrorCode.DEVICE_NOT_ALLOWED
    assert len(http.calls) == 1


@pytest.mark.asyncio
async def test_detected_apps_requires_explicit_tenant_inventory_opt_in() -> None:
    denied_http = FakeHttpClient()
    denied_client = service(denied_http)
    with pytest.raises(PolicyViolationError) as exc_info:
        await denied_client.list_detected_apps()
    assert exc_info.value.error.code == IntuneErrorCode.DETECTED_APPS_NOT_ALLOWED
    assert denied_http.calls == []

    allowed_http = FakeHttpClient(
        FakeResponse(
            200,
            {
                "value": [
                    {
                        "id": str(APP_ID),
                        "displayName": "Microsoft 365 Apps",
                        "platform": "windows",
                        "deviceCount": 2,
                    }
                ]
            },
        )
    )
    allowed_client = service(
        allowed_http,
        service_settings=settings(allow_tenant_detected_apps=True),
    )

    result = await allowed_client.list_detected_apps()

    assert result.value[0].id == APP_ID
    assert result.value[0].display_name == "Microsoft 365 Apps"
    assert allowed_http.calls[0][1].endswith("/deviceManagement/detectedApps")


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("action", "method_name", "suffix", "body", "destructive", "disruptive"),
    [
        (DeviceAction.SYNC_DEVICE, "sync_device", "syncDevice", None, False, False),
        (DeviceAction.REBOOT_NOW, "reboot_now", "rebootNow", None, True, True),
        (DeviceAction.REMOTE_LOCK, "remote_lock", "remoteLock", None, False, True),
        (DeviceAction.SHUT_DOWN, "shut_down", "shutDown", None, True, True),
        (
            DeviceAction.WINDOWS_DEFENDER_SCAN,
            "windows_defender_scan",
            "windowsDefenderScan",
            {"quickScan": True},
            False,
            False,
        ),
    ],
)
async def test_documented_actions_use_fixed_paths_and_audit_headers(
    action: DeviceAction,
    method_name: str,
    suffix: str,
    body: dict[str, bool] | None,
    destructive: bool,
    disruptive: bool,
) -> None:
    http = FakeHttpClient(FakeResponse(204, headers={"request-id": "graph-id"}))
    client = service(http)
    evidence = confirmation(action, destructive_acknowledged=destructive)
    method = getattr(client, method_name)

    if action == DeviceAction.WINDOWS_DEFENDER_SCAN:
        result = await method(DEVICE_ID, quick_scan=True, confirmation=evidence)
    else:
        result = await method(DEVICE_ID, confirmation=evidence)

    assert result.action == action
    assert result.destructive is destructive
    assert result.disruptive is disruptive
    assert result.graph_request_id == "graph-id"
    request_method, url, kwargs = http.calls[0]
    assert request_method == "POST"
    assert url == (
        f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/"
        f"{DEVICE_ID}/{suffix}"
    )
    assert kwargs["headers"]["client-request-id"] == str(evidence.correlation_id)
    assert kwargs["headers"]["return-client-request-id"] == "true"
    assert kwargs["headers"]["Idempotency-Key"] == str(evidence.idempotency_key)
    if body is None:
        assert "json" not in kwargs
        assert "Content-Type" not in kwargs["headers"]
    else:
        assert kwargs["json"] == body
        assert kwargs["headers"]["Content-Type"] == "application/json"


@pytest.mark.asyncio
async def test_actions_require_allowlists_and_bound_fresh_confirmation() -> None:
    denied_http = FakeHttpClient()
    denied_client = service(
        denied_http,
        service_settings=settings(allowed_actions=frozenset()),
    )
    with pytest.raises(PolicyViolationError) as exc_info:
        await denied_client.sync_device(
            DEVICE_ID, confirmation=confirmation(DeviceAction.SYNC_DEVICE)
        )
    assert exc_info.value.error.code == IntuneErrorCode.ACTION_NOT_ALLOWED

    client = service(FakeHttpClient())
    with pytest.raises(ConfirmationError):
        await client.sync_device(DEVICE_ID, confirmation=None)
    with pytest.raises(ConfirmationError):
        await client.sync_device(
            DEVICE_ID,
            confirmation=confirmation(
                DeviceAction.SYNC_DEVICE, device_id=SECOND_DEVICE_ID
            ),
        )
    with pytest.raises(ConfirmationError):
        await client.sync_device(
            DEVICE_ID,
            confirmation=confirmation(
                DeviceAction.SYNC_DEVICE,
                confirmed_at=NOW - timedelta(minutes=10),
                expires_at=NOW - timedelta(minutes=5),
            ),
        )
    assert denied_http.calls == []


@pytest.mark.asyncio
async def test_destructive_actions_require_separate_acknowledgement() -> None:
    client = service(FakeHttpClient())

    with pytest.raises(ConfirmationError) as exc_info:
        await client.reboot_now(
            DEVICE_ID, confirmation=confirmation(DeviceAction.REBOOT_NOW)
        )

    assert "not explicitly acknowledged" in exc_info.value.error.message


@pytest.mark.asyncio
async def test_idempotency_replays_same_action_and_rejects_key_reuse() -> None:
    http = FakeHttpClient(FakeResponse(204))
    client = service(http)
    evidence = confirmation(DeviceAction.SYNC_DEVICE)

    first = await client.sync_device(DEVICE_ID, confirmation=evidence)
    replay = await client.sync_device(DEVICE_ID, confirmation=evidence)

    assert first.replayed is False
    assert replay.replayed is True
    assert len(http.calls) == 1

    conflicting = confirmation(
        DeviceAction.REMOTE_LOCK,
        idempotency_key=evidence.idempotency_key,
        correlation_id=evidence.correlation_id,
    )
    with pytest.raises(IdempotencyConflictError) as exc_info:
        await client.remote_lock(DEVICE_ID, confirmation=conflicting)
    assert exc_info.value.error.code == IntuneErrorCode.IDEMPOTENCY_CONFLICT
    assert len(http.calls) == 1


@pytest.mark.asyncio
async def test_rotate_bitlocker_keys_is_stably_unsupported_and_never_calls_beta() -> (
    None
):
    http = FakeHttpClient()
    client = service(http)

    with pytest.raises(UnsupportedCapabilityError) as exc_info:
        await client.rotate_bitlocker_keys(
            DEVICE_ID,
            confirmation=confirmation(
                DeviceAction.ROTATE_BITLOCKER_KEYS,
                destructive_acknowledged=True,
            ),
        )

    assert exc_info.value.error.code == IntuneErrorCode.CAPABILITY_UNSUPPORTED
    assert "not documented for Microsoft Graph v1.0" in exc_info.value.error.message
    assert http.calls == []
    capability = {item.action: item for item in IntuneService.capabilities()}[
        DeviceAction.ROTATE_BITLOCKER_KEYS
    ]
    assert capability.supported is False
    assert capability.api_version is None


@pytest.mark.asyncio
async def test_graph_error_is_typed_and_preserves_correlation() -> None:
    http = FakeHttpClient(
        FakeResponse(
            429,
            {
                "error": {
                    "code": "TooManyRequests",
                    "message": "Slow down",
                    "innerError": {"request-id": "inner-request-id"},
                }
            },
        )
    )
    client = service(http)

    with pytest.raises(GraphRequestError) as exc_info:
        await client.get_managed_device(DEVICE_ID)

    error = exc_info.value.error
    assert error.code == IntuneErrorCode.GRAPH_REQUEST_FAILED
    assert error.http_status == 429
    assert error.graph_code == "TooManyRequests"
    assert error.graph_request_id == "inner-request-id"
    assert error.retryable is True
    assert error.correlation_id is not None


@pytest.mark.asyncio
async def test_non_graph_token_audience_fails_before_http() -> None:
    http = FakeHttpClient()
    tokens = FakeTokenProvider(
        {"access_token": "wrong-token", "audience": "https://example.com"}
    )
    client = service(http, tokens=tokens)

    with pytest.raises(AuthenticationError) as exc_info:
        await client.get_managed_device(DEVICE_ID)

    assert exc_info.value.error.code == IntuneErrorCode.AUTHENTICATION_FAILED
    assert http.calls == []
