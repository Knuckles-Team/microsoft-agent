"""Unit tests for the native Windows companion control-plane client."""

from __future__ import annotations

import asyncio
import json
from collections.abc import Mapping
from datetime import UTC, datetime, timedelta
from typing import Any
from uuid import UUID, uuid4

import pytest
from pydantic import ValidationError

from microsoft_agent.windows_companion import (
    CompanionActionKind,
    CompanionActionRequest,
    CompanionDevice,
    CompanionErrorCode,
    CompanionHttpResponse,
    ConfirmationEvidence,
    DeviceIdentity,
    FileReadAction,
    OfficeExportPdfAction,
    PowerAutomateDesktopRunAction,
    SystemInventoryAction,
    WindowsCompanionClient,
    WindowsCompanionClientError,
    WindowsCompanionSettings,
    WindowsServiceStartAction,
)

TENANT_ID = UUID("11111111-1111-4111-8111-111111111111")
ENTRA_DEVICE_ID = UUID("22222222-2222-4222-8222-222222222222")


def _identity(**overrides: Any) -> DeviceIdentity:
    values: dict[str, Any] = {
        "device_id": "sample-device-01",
        "tenant_id": TENANT_ID,
        "entra_device_id": ENTRA_DEVICE_ID,
        "certificate_thumbprint": "AB" * 20,
    }
    values.update(overrides)
    return DeviceIdentity(**values)


def _device(**overrides: Any) -> CompanionDevice:
    values: dict[str, Any] = {
        "identity": _identity(),
        "display_name": "Sample Device 01",
        "allowed_actions": {
            CompanionActionKind.SYSTEM_INVENTORY,
            CompanionActionKind.FILE_READ,
            CompanionActionKind.OFFICE_EXPORT_PDF,
            CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN,
            CompanionActionKind.WINDOWS_SERVICE_START,
        },
        "allowed_file_roots": (r"C:\Data\Documents",),
        "allowed_services": {"Spooler"},
        "allowed_desktop_flows": {"Publish presentation"},
    }
    values.update(overrides)
    return CompanionDevice(**values)


def _settings(**overrides: Any) -> WindowsCompanionSettings:
    values: dict[str, Any] = {
        "control_plane_url": "https://windows-relay.example.com",
        "token_audience": "https://windows-relay.example.com",
        "devices": {"primary": _device()},
    }
    values.update(overrides)
    return WindowsCompanionSettings(**values)


def _response(
    payload: Any = None,
    *,
    status: int = 200,
    headers: Mapping[str, str] | None = None,
) -> CompanionHttpResponse:
    body = b"" if payload is None else json.dumps(payload).encode()
    return CompanionHttpResponse(
        status_code=status, headers=dict(headers or {}), body=body
    )


class FakeTransport:
    def __init__(self, *responses: CompanionHttpResponse | Exception) -> None:
        self.responses = list(responses)
        self.calls: list[dict[str, Any]] = []

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
        self.calls.append(
            {
                "method": method,
                "url": url,
                "headers": dict(headers),
                "params": params,
                "body": body,
                "timeout": timeout,
            }
        )
        result = self.responses.pop(0)
        if isinstance(result, Exception):
            raise result
        return result


class FakeTokenProvider:
    def __init__(self, token: str = "companion-token", *, delay: float = 0) -> None:
        self.token = token
        self.delay = delay
        self.audiences: list[str] = []

    async def get_token(self, audience: str) -> str:
        self.audiences.append(audience)
        if self.delay:
            await asyncio.sleep(self.delay)
        return self.token


def _confirmation(kind: CompanionActionKind) -> ConfirmationEvidence:
    now = datetime.now(UTC)
    return ConfirmationEvidence(
        action_kind=kind,
        confirmed_by="user@example.com",
        confirmed_at=now,
        expires_at=now + timedelta(minutes=5),
        purpose="User approved this device action",
        authorization_reference="conversation/turn/42",
    )


def test_settings_require_https_outbound_relay() -> None:
    with pytest.raises(ValidationError):
        _settings(control_plane_url="http://localhost:8080")
    with pytest.raises(ValidationError):
        _settings(connection_mode="direct")


def test_unknown_shell_action_cannot_be_modeled() -> None:
    with pytest.raises(ValidationError):
        CompanionActionRequest.model_validate(
            {"action": {"kind": "shell.run", "command": "whoami"}}
        )


@pytest.mark.asyncio
async def test_health_uses_authenticated_per_device_endpoint() -> None:
    identity = _identity()
    payload = {
        "identity": identity.model_dump(mode="json"),
        "status": "online",
        "authenticated": True,
        "outbound_connected": True,
        "last_seen_at": "2026-07-14T12:00:00Z",
        "companion_version": "1.0.0",
        "capabilities": ["system.inventory"],
    }
    transport = FakeTransport(_response(payload))
    tokens = FakeTokenProvider()
    client = WindowsCompanionClient(_settings(), tokens, transport)

    health = await client.get_health("primary")

    assert health.identity == identity
    assert health.outbound_connected is True
    call = transport.calls[0]
    assert call["url"] == (
        "https://windows-relay.example.com/v1/devices/sample-device-01/health"
    )
    assert call["headers"]["Authorization"] == "Bearer companion-token"
    assert call["headers"]["X-Microsoft-Agent-Device-ID"] == "sample-device-01"
    assert tokens.audiences == ["https://windows-relay.example.com"]


@pytest.mark.asyncio
async def test_health_rejects_mismatched_device_identity() -> None:
    payload = {
        "identity": _identity(device_id="other-device").model_dump(mode="json"),
        "status": "online",
        "authenticated": True,
        "outbound_connected": True,
        "last_seen_at": "2026-07-14T12:00:00Z",
        "companion_version": "1.0.0",
    }
    client = WindowsCompanionClient(
        _settings(), FakeTokenProvider(), FakeTransport(_response(payload))
    )

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.get_health("primary")

    assert exc_info.value.error.code is CompanionErrorCode.INVALID_RESPONSE


@pytest.mark.asyncio
async def test_read_only_inventory_can_be_submitted_without_confirmation() -> None:
    request = CompanionActionRequest(action=SystemInventoryAction())
    receipt = {
        "action_id": str(request.action_id),
        "device_id": "sample-device-01",
        "status": "accepted",
        "accepted_at": "2026-07-14T12:00:00Z",
        "status_url": "/v1/actions/status",
    }
    transport = FakeTransport(_response(receipt, status=202))
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), transport)

    result = await client.submit_action("primary", request)

    assert result.action_id == request.action_id
    call = transport.calls[0]
    body = json.loads(call["body"])
    assert body["request"]["action"]["kind"] == "system.inventory"
    assert body["policy"]["confirmation"] == "none"
    assert body["expected_device_identity"]["entra_device_id"] == str(ENTRA_DEVICE_ID)
    assert call["headers"]["Idempotency-Key"] == request.idempotency_key


@pytest.mark.asyncio
async def test_sensitive_file_read_requires_confirmation() -> None:
    action = FileReadAction(path=r"C:\Data\Documents\notes.txt")
    request = CompanionActionRequest(action=action)
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), FakeTransport())

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.submit_action("primary", request)

    assert exc_info.value.error.code is CompanionErrorCode.POLICY
    assert "requires confirmation" in exc_info.value.error.message


@pytest.mark.asyncio
async def test_confirmed_file_read_inside_root_is_accepted() -> None:
    action = FileReadAction(path=r"C:\Data\Documents\notes.txt")
    request = CompanionActionRequest(
        action=action,
        confirmation=_confirmation(CompanionActionKind.FILE_READ),
    )
    receipt = {
        "action_id": str(request.action_id),
        "device_id": "sample-device-01",
        "status": "accepted",
        "accepted_at": "2026-07-14T12:00:00Z",
    }
    transport = FakeTransport(_response(receipt, status=202))
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), transport)

    await client.submit_action("primary", request)

    assert len(transport.calls) == 1


@pytest.mark.asyncio
async def test_file_path_traversal_outside_root_is_rejected() -> None:
    action = FileReadAction(path=r"C:\Data\Documents\..\Secrets\passwords.txt")
    request = CompanionActionRequest(
        action=action,
        confirmation=_confirmation(CompanionActionKind.FILE_READ),
    )
    transport = FakeTransport()
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), transport)

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.submit_action("primary", request)

    assert exc_info.value.error.code is CompanionErrorCode.POLICY
    assert "outside" in exc_info.value.error.message
    assert not transport.calls


@pytest.mark.asyncio
async def test_office_export_checks_both_source_and_output_roots() -> None:
    action = OfficeExportPdfAction(
        application="powerpoint",
        source_path=r"C:\Data\Documents\deck.pptx",
        output_path=r"D:\Public\deck.pdf",
    )
    request = CompanionActionRequest(
        action=action,
        confirmation=_confirmation(CompanionActionKind.OFFICE_EXPORT_PDF),
    )
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), FakeTransport())

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.submit_action("primary", request)

    assert exc_info.value.error.code is CompanionErrorCode.POLICY


@pytest.mark.asyncio
async def test_windows_service_name_must_be_allowlisted() -> None:
    action = WindowsServiceStartAction(service_name="RemoteRegistry")
    request = CompanionActionRequest(
        action=action,
        confirmation=_confirmation(CompanionActionKind.WINDOWS_SERVICE_START),
    )
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), FakeTransport())

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.submit_action("primary", request)

    assert "service is not allowlisted" in exc_info.value.error.message


@pytest.mark.asyncio
async def test_desktop_flow_name_must_be_allowlisted() -> None:
    action = PowerAutomateDesktopRunAction(flow_name="Unknown desktop flow")
    request = CompanionActionRequest(
        action=action,
        confirmation=_confirmation(CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN),
    )
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), FakeTransport())

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.submit_action("primary", request)

    assert "Desktop flow is not allowlisted" in exc_info.value.error.message


@pytest.mark.asyncio
async def test_confirmation_scope_and_expiry_are_enforced() -> None:
    now = datetime.now(UTC)
    wrong_scope = ConfirmationEvidence(
        action_kind=CompanionActionKind.FILE_WRITE,
        confirmed_by="user@example.com",
        confirmed_at=now - timedelta(minutes=2),
        expires_at=now + timedelta(minutes=2),
        purpose="Approve a different action",
        authorization_reference="conversation/turn/40",
    )
    action = FileReadAction(path=r"C:\Data\Documents\notes.txt")
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), FakeTransport())

    with pytest.raises(WindowsCompanionClientError) as wrong_scope_error:
        await client.submit_action(
            "primary", CompanionActionRequest(action=action, confirmation=wrong_scope)
        )
    assert "different action" in wrong_scope_error.value.error.message

    expired = ConfirmationEvidence(
        action_kind=CompanionActionKind.FILE_READ,
        confirmed_by="user@example.com",
        confirmed_at=now - timedelta(minutes=2),
        expires_at=now - timedelta(minutes=1),
        purpose="Expired approval",
        authorization_reference="conversation/turn/41",
    )
    with pytest.raises(WindowsCompanionClientError) as expired_error:
        await client.submit_action(
            "primary", CompanionActionRequest(action=action, confirmation=expired)
        )
    assert "expired" in expired_error.value.error.message


@pytest.mark.asyncio
async def test_unknown_device_alias_is_rejected_without_network_access() -> None:
    transport = FakeTransport()
    tokens = FakeTokenProvider()
    client = WindowsCompanionClient(_settings(), tokens, transport)

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.get_health("unconfigured-laptop")

    assert exc_info.value.error.code is CompanionErrorCode.POLICY
    assert not transport.calls
    assert not tokens.audiences


@pytest.mark.asyncio
async def test_offline_error_is_safely_shaped() -> None:
    transport = FakeTransport(
        _response(
            {"error": {"code": "device_disconnected", "message": "Laptop offline"}},
            status=503,
            headers={"Retry-After": "30", "x-ms-request-id": "relay-123"},
        )
    )
    client = WindowsCompanionClient(_settings(), FakeTokenProvider(), transport)

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.get_health("primary")

    error = exc_info.value.error
    assert error.code is CompanionErrorCode.OFFLINE
    assert error.upstream_code == "device_disconnected"
    assert error.message == "Laptop offline"
    assert error.retry_after_seconds == 30
    assert error.correlation_id == "relay-123"


@pytest.mark.asyncio
async def test_token_timeout_is_normalized() -> None:
    client = WindowsCompanionClient(
        _settings(timeout_seconds=0.01),
        FakeTokenProvider(delay=0.1),
        FakeTransport(),
    )

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.get_health("primary")

    assert exc_info.value.error.code is CompanionErrorCode.TIMEOUT


@pytest.mark.asyncio
async def test_action_result_must_match_requested_action_and_device() -> None:
    action_id = uuid4()
    payload = {
        "action_id": str(action_id),
        "device_id": "another-device",
        "status": "succeeded",
        "output": {"ok": True},
    }
    client = WindowsCompanionClient(
        _settings(), FakeTokenProvider(), FakeTransport(_response(payload))
    )

    with pytest.raises(WindowsCompanionClientError) as exc_info:
        await client.get_action_result("primary", action_id)

    assert exc_info.value.error.code is CompanionErrorCode.INVALID_RESPONSE
