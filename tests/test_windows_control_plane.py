"""Tests for the durable, authenticated Windows companion control plane."""

from __future__ import annotations

import json
import os
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any
from uuid import UUID

import httpx
import pytest
from pydantic import ValidationError

from microsoft_agent.windows_companion import (
    CompanionActionFailure,
    CompanionActionKind,
    CompanionActionRequest,
    CompanionActionResult,
    CompanionActionStatus,
    CompanionDevice,
    ConfirmationEvidence,
    DeviceIdentity,
    FileReadAction,
    SystemInventoryAction,
    WindowsServiceStartAction,
)
from microsoft_agent.windows_companion_service import (
    CompanionCertificateSettings,
    NativeAdapterSettings,
    WindowsCompanionServiceConfig,
    build_worker,
    load_service_config,
    validate_local_service_config,
)
from microsoft_agent.windows_companion_service import (
    main as service_main,
)
from microsoft_agent.windows_control_plane import (
    ActionSubmission,
    AuthenticatedPrincipal,
    ControlPlaneDeviceRegistration,
    EntraJwtValidatorSettings,
    HttpOutboundRelayTransport,
    SQLiteCompanionStore,
    StaticTokenValidator,
    WindowsControlPlaneLimits,
    WindowsControlPlaneSettings,
    create_windows_control_plane_app,
)
from microsoft_agent.windows_runtime import RelayPollBatch, WindowsRuntimeError

TENANT_ID = UUID("11111111-1111-1111-1111-111111111111")
ENTRA_DEVICE_ID = UUID("22222222-2222-2222-2222-222222222222")
DEVICE_ID = "laptop-1"
THUMBPRINT = "A" * 40
TOKEN_AUDIENCE = "00000000-0000-4000-8000-000000000010"
TOKEN_RESOURCE = f"api://{TOKEN_AUDIENCE}"


def _identity(device_id: str = DEVICE_ID) -> DeviceIdentity:
    return DeviceIdentity(
        device_id=device_id,
        tenant_id=TENANT_ID,
        entra_device_id=ENTRA_DEVICE_ID,
        certificate_thumbprint=THUMBPRINT,
    )


def _device() -> CompanionDevice:
    return CompanionDevice(
        identity=_identity(),
        display_name="Test laptop",
        allowed_actions=frozenset(
            {
                CompanionActionKind.SYSTEM_INVENTORY,
                CompanionActionKind.FILE_READ,
                CompanionActionKind.WINDOWS_SERVICE_START,
            }
        ),
        allowed_file_roots=(r"C:\Allowed",),
        allowed_services=frozenset({"Spooler"}),
    )


def _settings(**changes: Any) -> WindowsControlPlaneSettings:
    values: dict[str, Any] = {
        "token_audience": TOKEN_AUDIENCE,
        "devices": {
            DEVICE_ID: ControlPlaneDeviceRegistration(
                device=_device(), device_principal_subjects=frozenset({"device-app"})
            )
        },
        "trusted_proxy_mtls_header": "X-Client-Cert-Thumbprint",
        "trusted_proxy_client_hosts": frozenset({"127.0.0.1"}),
        "require_device_mtls": True,
    }
    values.update(changes)
    return WindowsControlPlaneSettings(**values)


def _principal(
    subject: str,
    *,
    controller: bool = False,
    device_claim: UUID | None = None,
    audience: str = TOKEN_AUDIENCE,
    tenant_id: UUID = TENANT_ID,
) -> AuthenticatedPrincipal:
    return AuthenticatedPrincipal(
        tenant_id=tenant_id,
        subject=subject,
        audience=audience,
        issuer=f"https://login.microsoftonline.com/{tenant_id}/v2.0",
        entra_device_id=device_claim,
        roles=frozenset(
            {"WindowsCompanion.Control" if controller else "WindowsCompanion.Device"}
        ),
    )


def _validator() -> StaticTokenValidator:
    return StaticTokenValidator(
        {
            "controller-token": _principal("controller", controller=True),
            "device-token": _principal("device-app", device_claim=ENTRA_DEVICE_ID),
            "wrong-audience": _principal(
                "controller", controller=True, audience="api://other"
            ),
            "wrong-tenant": _principal(
                "controller",
                controller=True,
                tenant_id=UUID("33333333-3333-3333-3333-333333333333"),
            ),
        }
    )


def _controller_headers() -> dict[str, str]:
    return {
        "Authorization": "Bearer controller-token",
        "X-Microsoft-Agent-Device-ID": DEVICE_ID,
    }


def _device_headers(*, certificate: str = THUMBPRINT) -> dict[str, str]:
    return {
        "Authorization": "Bearer device-token",
        "X-Microsoft-Agent-Device-ID": DEVICE_ID,
        "X-Client-Cert-Thumbprint": certificate,
    }


def _inventory_request(key: str = "request-1") -> CompanionActionRequest:
    now = datetime.now(UTC)
    return CompanionActionRequest(
        action=SystemInventoryAction(),
        requested_at=now,
        expires_at=now + timedelta(minutes=5),
        idempotency_key=key,
    )


def _submission(
    settings: WindowsControlPlaneSettings,
    request: CompanionActionRequest | None = None,
    *,
    identity: DeviceIdentity | None = None,
) -> ActionSubmission:
    action_request = request or _inventory_request()
    return ActionSubmission(
        request=action_request,
        policy=settings.action_policies[
            CompanionActionKind(action_request.action.kind)
        ],
        expected_device_identity=identity or _identity(),
    )


@pytest.fixture
def control_plane(
    tmp_path: Path,
) -> tuple[Any, SQLiteCompanionStore, WindowsControlPlaneSettings]:
    settings = _settings()
    store = SQLiteCompanionStore(tmp_path / "queue.db")
    app = create_windows_control_plane_app(settings, store, _validator())
    return app, store, settings


@pytest.mark.asyncio
async def test_controller_action_round_trip_is_durable(control_plane: Any) -> None:
    app, store, settings = control_plane
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        submission = _submission(settings)
        created = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=submission.model_dump(mode="json"),
        )
        assert created.status_code == 202
        receipt = created.json()
        action_id = receipt["action_id"]

        polled = await client.post(
            f"/v1/devices/{DEVICE_ID}/relay/poll",
            headers=_device_headers(),
            json={
                "maximum_actions": 10,
                "wait_seconds": 0,
                "companion_version": "1.0.0",
                "capabilities": [CompanionActionKind.SYSTEM_INVENTORY.value],
            },
        )
        assert polled.status_code == 200
        delivery = polled.json()["deliveries"][0]
        assert delivery["request"]["action_id"] == action_id

        running = await client.get(
            f"/v1/devices/{DEVICE_ID}/actions/{action_id}",
            headers=_controller_headers(),
        )
        assert running.json()["status"] == CompanionActionStatus.RUNNING.value

        result = CompanionActionResult(
            action_id=UUID(action_id),
            device_id=DEVICE_ID,
            status=CompanionActionStatus.SUCCEEDED,
            started_at=datetime.now(UTC),
            completed_at=datetime.now(UTC),
            output={"hostname": "test-laptop"},
        )
        acknowledged = await client.post(
            (f"/v1/devices/{DEVICE_ID}/relay/actions/{delivery['delivery_id']}/ack"),
            headers=_device_headers(),
            json=result.model_dump(mode="json"),
        )
        assert acknowledged.status_code == 204

        completed = await client.get(
            f"/v1/devices/{DEVICE_ID}/actions/{action_id}",
            headers=_controller_headers(),
        )
        assert completed.status_code == 200
        assert completed.json()["output"] == {"hostname": "test-laptop"}

    reopened = SQLiteCompanionStore(store.path)
    durable = await reopened.get_action(DEVICE_ID, UUID(action_id))
    assert durable.status is CompanionActionStatus.SUCCEEDED


@pytest.mark.asyncio
async def test_submission_idempotency_and_conflict(control_plane: Any) -> None:
    app, _, settings = control_plane
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        submission = _submission(settings)
        first = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=submission.model_dump(mode="json"),
        )
        replay = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=submission.model_dump(mode="json"),
        )
        conflict_request = _inventory_request("request-1")
        conflict = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=_submission(settings, conflict_request).model_dump(mode="json"),
        )

    assert first.status_code == 202
    assert replay.status_code == 202
    assert replay.json()["action_id"] == first.json()["action_id"]
    assert conflict.status_code == 409


@pytest.mark.asyncio
async def test_bearer_audience_tenant_and_role_are_enforced(control_plane: Any) -> None:
    app, _, _ = control_plane
    transport = httpx.ASGITransport(app=app)
    path = f"/v1/devices/{DEVICE_ID}/health"
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        missing = await client.get(path)
        bad_audience = await client.get(
            path, headers={"Authorization": "Bearer wrong-audience"}
        )
        wrong_tenant = await client.get(
            path, headers={"Authorization": "Bearer wrong-tenant"}
        )
        device_as_controller = await client.get(
            path, headers={"Authorization": "Bearer device-token"}
        )

    assert missing.status_code == 401
    assert bad_audience.status_code == 401
    assert wrong_tenant.status_code == 403
    assert device_as_controller.status_code == 403


@pytest.mark.asyncio
async def test_request_size_and_validation_errors_are_safely_bounded(
    tmp_path: Path,
) -> None:
    settings = _settings()
    limits = WindowsControlPlaneLimits(
        maximum_submission_bytes=1024,
        maximum_result_bytes=1024,
    )
    store = SQLiteCompanionStore(tmp_path / "small.db", limits)
    app = create_windows_control_plane_app(settings, store, _validator())
    transport = httpx.ASGITransport(app=app)
    path = f"/v1/devices/{DEVICE_ID}/actions"
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        oversized = await client.post(
            path,
            headers=_controller_headers(),
            content=b"x" * (limits.maximum_http_request_bytes + 1),
        )
        malformed = await client.post(
            path,
            headers=_controller_headers(),
            json={"request": {"private_value": "must-not-be-echoed"}},
        )

    assert oversized.status_code == 413
    assert malformed.status_code == 422
    assert malformed.json() == {"detail": "Request body or path parameters are invalid"}
    assert "must-not-be-echoed" not in malformed.text


@pytest.mark.asyncio
async def test_device_poll_requires_token_device_header_and_mtls(
    control_plane: Any,
) -> None:
    app, _, _ = control_plane
    path = f"/v1/devices/{DEVICE_ID}/relay/poll"
    body = {"maximum_actions": 1, "wait_seconds": 0}
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        controller = await client.post(path, headers=_controller_headers(), json=body)
        no_certificate_headers = _device_headers()
        no_certificate_headers.pop("X-Client-Cert-Thumbprint")
        no_certificate = await client.post(
            path, headers=no_certificate_headers, json=body
        )
        wrong_certificate = await client.post(
            path, headers=_device_headers(certificate="B" * 40), json=body
        )
        valid = await client.post(path, headers=_device_headers(), json=body)

    assert controller.status_code == 403
    assert no_certificate.status_code == 403
    assert wrong_certificate.status_code == 403
    assert valid.status_code == 200


@pytest.mark.asyncio
async def test_payload_identity_cannot_override_configured_identity(
    control_plane: Any,
) -> None:
    app, _, settings = control_plane
    wrong = DeviceIdentity(
        device_id=DEVICE_ID,
        tenant_id=TENANT_ID,
        entra_device_id=UUID("44444444-4444-4444-4444-444444444444"),
        certificate_thumbprint="B" * 40,
    )
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        response = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=_submission(settings, identity=wrong).model_dump(mode="json"),
        )

    assert response.status_code == 403


@pytest.mark.asyncio
async def test_path_service_and_confirmation_policy_are_rechecked(
    control_plane: Any,
) -> None:
    app, _, settings = control_plane
    now = datetime.now(UTC)
    outside_action = FileReadAction(path=r"C:\Outside\secret.txt")
    outside_request = CompanionActionRequest(
        action=outside_action,
        requested_at=now,
        expires_at=now + timedelta(minutes=5),
        confirmation=ConfirmationEvidence(
            action_kind=CompanionActionKind.FILE_READ,
            confirmed_by="user",
            confirmed_at=now - timedelta(seconds=1),
            expires_at=now + timedelta(minutes=2),
            purpose="Read requested file",
            authorization_reference="approval/1",
        ),
    )
    service_request = CompanionActionRequest(
        action=WindowsServiceStartAction(service_name="Unlisted"),
        requested_at=now,
        expires_at=now + timedelta(minutes=5),
        confirmation=ConfirmationEvidence(
            action_kind=CompanionActionKind.WINDOWS_SERVICE_START,
            confirmed_by="user",
            confirmed_at=now - timedelta(seconds=1),
            expires_at=now + timedelta(minutes=2),
            purpose="Start service",
            authorization_reference="approval/2",
        ),
    )
    no_confirmation = CompanionActionRequest(
        action=FileReadAction(path=r"C:\Allowed\secret.txt"),
        requested_at=now,
        expires_at=now + timedelta(minutes=5),
    )
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        responses = [
            await client.post(
                f"/v1/devices/{DEVICE_ID}/actions",
                headers=_controller_headers(),
                json=_submission(settings, request).model_dump(mode="json"),
            )
            for request in (outside_request, service_request, no_confirmation)
        ]

    assert [item.status_code for item in responses] == [403, 403, 403]


@pytest.mark.asyncio
async def test_health_reflects_authenticated_outbound_poll(control_plane: Any) -> None:
    app, _, _ = control_plane
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        before = await client.get(
            f"/v1/devices/{DEVICE_ID}/health", headers=_controller_headers()
        )
        await client.post(
            f"/v1/devices/{DEVICE_ID}/relay/poll",
            headers=_device_headers(),
            json={
                "wait_seconds": 0,
                "companion_version": "2.0.0",
                "capabilities": [CompanionActionKind.SYSTEM_INVENTORY.value],
            },
        )
        after = await client.get(
            f"/v1/devices/{DEVICE_ID}/health", headers=_controller_headers()
        )

    assert before.json()["status"] == "offline"
    assert before.json()["authenticated"] is False
    assert after.json()["status"] == "online"
    assert after.json()["authenticated"] is True
    assert after.json()["companion_version"] == "2.0.0"


@pytest.mark.asyncio
async def test_acknowledgement_is_idempotent_but_conflicts_on_changed_result(
    control_plane: Any,
) -> None:
    app, _, settings = control_plane
    transport = httpx.ASGITransport(app=app)
    async with httpx.AsyncClient(
        transport=transport, base_url="https://control.test"
    ) as client:
        submitted = await client.post(
            f"/v1/devices/{DEVICE_ID}/actions",
            headers=_controller_headers(),
            json=_submission(settings).model_dump(mode="json"),
        )
        action_id = UUID(submitted.json()["action_id"])
        poll = await client.post(
            f"/v1/devices/{DEVICE_ID}/relay/poll",
            headers=_device_headers(),
            json={"wait_seconds": 0},
        )
        delivery_id = poll.json()["deliveries"][0]["delivery_id"]
        result = CompanionActionResult(
            action_id=action_id,
            device_id=DEVICE_ID,
            status=CompanionActionStatus.FAILED,
            completed_at=datetime.now(UTC),
            error=CompanionActionFailure(code="test", message="failed"),
        )
        path = f"/v1/devices/{DEVICE_ID}/relay/actions/{delivery_id}/ack"
        first = await client.post(
            path, headers=_device_headers(), json=result.model_dump(mode="json")
        )
        replay = await client.post(
            path, headers=_device_headers(), json=result.model_dump(mode="json")
        )
        changed = result.model_copy(
            update={
                "error": CompanionActionFailure(code="different", message="changed")
            }
        )
        conflict = await client.post(
            path, headers=_device_headers(), json=changed.model_dump(mode="json")
        )

    assert first.status_code == 204
    assert replay.status_code == 204
    assert conflict.status_code == 409


@pytest.mark.asyncio
async def test_store_expires_unclaimed_actions(tmp_path: Path) -> None:
    store = SQLiteCompanionStore(tmp_path / "expired.db")
    settings = _settings(require_device_mtls=False, trusted_proxy_mtls_header=None)
    now = datetime.now(UTC)
    request = CompanionActionRequest(
        action=SystemInventoryAction(),
        requested_at=now - timedelta(minutes=2),
        expires_at=now - timedelta(minutes=1),
    )
    # Direct store insertion is intentionally allowed only after route policy;
    # it lets this test exercise durable expiry independently of wall-clock races.
    receipt = await store.enqueue(DEVICE_ID, _submission(settings, request))

    result = await store.get_action(DEVICE_ID, receipt.action_id)

    assert result.status is CompanionActionStatus.EXPIRED
    assert result.error and result.error.code == "expired"


@pytest.mark.asyncio
async def test_store_enforces_pending_capacity(tmp_path: Path) -> None:
    limits = WindowsControlPlaneLimits(maximum_pending_per_device=1)
    store = SQLiteCompanionStore(tmp_path / "bounded.db", limits)
    settings = _settings(require_device_mtls=False, trusted_proxy_mtls_header=None)
    await store.enqueue(DEVICE_ID, _submission(settings, _inventory_request("first")))

    with pytest.raises(RuntimeError, match="queue is full"):
        await store.enqueue(
            DEVICE_ID, _submission(settings, _inventory_request("second"))
        )


def test_mtls_configuration_is_explicit() -> None:
    with pytest.raises(ValidationError, match="trusted proxy"):
        _settings(trusted_proxy_mtls_header=None, require_device_mtls=True)


def test_entra_validator_settings_reject_weak_algorithms_and_bad_jwks() -> None:
    with pytest.raises(ValidationError, match="RSA SHA-2"):
        EntraJwtValidatorSettings(
            tenant_id=TENANT_ID,
            audience=TOKEN_AUDIENCE,
            jwks={"keys": [{"kid": "one", "kty": "RSA", "n": "x", "e": "AQAB"}]},
            allowed_algorithms=("HS256",),
        )
    with pytest.raises(ValidationError, match="signing key"):
        EntraJwtValidatorSettings(
            tenant_id=TENANT_ID, audience=TOKEN_AUDIENCE, jwks={"keys": []}
        )


class FakeTokenProvider:
    async def get_token(self, audience: str) -> str:
        assert audience == TOKEN_RESOURCE
        return "device-access-token"


class FakeHttpResponse:
    def __init__(self, status: int, body: bytes = b"") -> None:
        self.status = status
        self._body = body

    def read(self, maximum_bytes: int = -1) -> bytes:
        return self._body if maximum_bytes < 0 else self._body[:maximum_bytes]

    def __enter__(self) -> FakeHttpResponse:
        return self

    def __exit__(self, *args: Any) -> None:
        return None


class FakeOpener:
    def __init__(self, responses: list[FakeHttpResponse]) -> None:
        self.responses = responses
        self.requests: list[Any] = []

    def open(self, request: Any, timeout: float) -> FakeHttpResponse:
        assert timeout == 5
        self.requests.append(request)
        return self.responses.pop(0)


@pytest.mark.asyncio
async def test_http_outbound_transport_serializes_poll_and_ack_without_network() -> (
    None
):
    poll_body = RelayPollBatch(cursor="next", deliveries=()).model_dump_json().encode()
    fake = FakeOpener([FakeHttpResponse(200, poll_body), FakeHttpResponse(204)])
    transport = HttpOutboundRelayTransport(
        "https://control.example",
        TOKEN_RESOURCE,
        FakeTokenProvider(),
        timeout_seconds=5,
        companion_version="test",
        capabilities=frozenset({CompanionActionKind.SYSTEM_INVENTORY}),
    )
    transport._opener = fake  # type: ignore[assignment]

    batch = await transport.poll(
        _identity(), cursor=None, maximum_actions=3, wait_seconds=0
    )
    result = CompanionActionResult(
        action_id=UUID("55555555-5555-5555-5555-555555555555"),
        device_id=DEVICE_ID,
        status=CompanionActionStatus.SUCCEEDED,
        completed_at=datetime.now(UTC),
        output={},
    )
    await transport.acknowledge(_identity(), "delivery-1", result)

    assert batch.cursor == "next"
    assert len(fake.requests) == 2
    assert fake.requests[0].full_url.endswith(f"/v1/devices/{DEVICE_ID}/relay/poll")
    assert fake.requests[0].headers["Authorization"] == "Bearer device-access-token"
    poll_payload = json.loads(fake.requests[0].data)
    assert poll_payload["maximum_actions"] == 3
    assert poll_payload["capabilities"] == [CompanionActionKind.SYSTEM_INVENTORY.value]
    assert fake.requests[1].full_url.endswith("/delivery-1/ack")


@pytest.mark.asyncio
async def test_http_outbound_transport_normalizes_http_failures() -> None:
    fake = FakeOpener([FakeHttpResponse(403, b"forbidden")])
    transport = HttpOutboundRelayTransport(
        "https://control.example",
        TOKEN_RESOURCE,
        FakeTokenProvider(),
        timeout_seconds=5,
    )
    transport._opener = fake  # type: ignore[assignment]

    with pytest.raises(WindowsRuntimeError, match="HTTP 403"):
        await transport.poll(
            _identity(), cursor=None, maximum_actions=1, wait_seconds=0
        )


def _service_certificate(tmp_path: Path) -> tuple[Path, Path, str]:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import NameOID

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = issuer = x509.Name(
        [x509.NameAttribute(NameOID.COMMON_NAME, "windows-companion-test")]
    )
    now = datetime.now(UTC)
    certificate = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(minutes=1))
        .not_valid_after(now + timedelta(days=1))
        .sign(key, hashes.SHA256())
    )
    certificate_path = tmp_path / "device-cert.pem"
    key_path = tmp_path / "device-key.pem"
    certificate_path.write_bytes(certificate.public_bytes(serialization.Encoding.PEM))
    key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption(),
        )
    )
    key_path.chmod(0o600)
    thumbprint = certificate.fingerprint(hashes.SHA1()).hex().upper()  # noqa: S324
    return certificate_path, key_path, thumbprint


def _service_config(tmp_path: Path) -> WindowsCompanionServiceConfig:
    certificate_path, key_path, thumbprint = _service_certificate(tmp_path)
    identity = DeviceIdentity(
        device_id=DEVICE_ID,
        tenant_id=TENANT_ID,
        entra_device_id=ENTRA_DEVICE_ID,
        certificate_thumbprint=thumbprint,
    )
    device = CompanionDevice(
        identity=identity,
        display_name="Service test laptop",
        allowed_actions=frozenset({CompanionActionKind.SYSTEM_INVENTORY}),
        allowed_file_roots=(r"C:\Allowed",),
    )
    return WindowsCompanionServiceConfig(
        control_plane_url="https://control.example",
        token_audience=TOKEN_RESOURCE,
        tenant_id=TENANT_ID,
        client_id=UUID("66666666-6666-6666-6666-666666666666"),
        device=device,
        file_root_bindings={r"C:\Allowed": tmp_path},
        certificate=CompanionCertificateSettings(
            thumbprint=thumbprint,
            client_certificate_path=certificate_path,
            client_private_key_path=key_path,
        ),
        adapters=NativeAdapterSettings(),
        companion_version="test",
    )


def test_service_config_load_and_validation_are_functional(
    tmp_path: Path, capsys: Any
) -> None:
    config = _service_config(tmp_path)
    config_path = tmp_path / "windows-companion.json"
    config_path.write_text(config.model_dump_json(indent=2), encoding="utf-8")

    loaded = load_service_config(config_path)
    validate_local_service_config(loaded)
    exit_code = service_main(["--config", os.fspath(config_path), "--validate-config"])

    assert loaded.device.identity == config.device.identity
    assert exit_code == 0
    assert "valid and enabled" in capsys.readouterr().out


class EmptyRelay:
    async def poll(
        self,
        identity: DeviceIdentity,
        *,
        cursor: str | None,
        maximum_actions: int,
        wait_seconds: float,
    ) -> RelayPollBatch:
        return RelayPollBatch(cursor=cursor)

    async def acknowledge(
        self, identity: DeviceIdentity, delivery_id: str, result: Any
    ) -> None:
        raise AssertionError("No delivery should be acknowledged")


def test_service_worker_can_be_built_with_injected_relay(tmp_path: Path) -> None:
    config = _service_config(tmp_path)

    worker = build_worker(config, relay_transport=EmptyRelay())

    assert worker.cursor is None


def test_service_fails_closed_when_allowed_adapter_is_disabled(tmp_path: Path) -> None:
    config = _service_config(tmp_path)
    device = config.device.model_copy(
        update={
            "allowed_actions": frozenset(
                {
                    CompanionActionKind.SYSTEM_INVENTORY,
                    CompanionActionKind.WINDOWS_SERVICE_START,
                }
            ),
            "allowed_services": frozenset({"Spooler"}),
        }
    )
    unsafe = config.model_copy(update={"device": device})

    with pytest.raises(ValueError, match="services adapter"):
        build_worker(unsafe, relay_transport=EmptyRelay())


def test_service_config_rejects_secret_or_unknown_fields(tmp_path: Path) -> None:
    config = _service_config(tmp_path).model_dump(mode="json")
    config["client_secret"] = "must-never-be-supported"

    with pytest.raises(ValidationError, match="client_secret"):
        WindowsCompanionServiceConfig.model_validate(config)


def test_disabled_bootstrap_config_does_not_require_credentials(
    tmp_path: Path, capsys: Any
) -> None:
    config = _service_config(tmp_path)
    disabled = config.model_copy(
        update={
            "device": config.device.model_copy(update={"enabled": False}),
            "certificate": config.certificate.model_copy(
                update={
                    "client_certificate_path": tmp_path / "not-provisioned-cert.pem",
                    "client_private_key_path": tmp_path / "not-provisioned-key.pem",
                }
            ),
        }
    )
    path = tmp_path / "disabled.json"
    path.write_text(disabled.model_dump_json(indent=2), encoding="utf-8")

    exit_code = service_main(["--config", os.fspath(path), "--validate-config"])

    assert exit_code == 0
    assert "device is disabled" in capsys.readouterr().out
    with pytest.raises(ValueError, match="disabled"):
        build_worker(disabled, relay_transport=EmptyRelay())
