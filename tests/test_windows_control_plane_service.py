"""Tests for the runnable Windows companion control-plane host."""

from __future__ import annotations

import json
import os
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any
from uuid import UUID

import jwt
import pytest
from cryptography.hazmat.primitives.asymmetric import rsa
from pydantic import ValidationError

from microsoft_agent.windows_companion import CompanionDevice, DeviceIdentity
from microsoft_agent.windows_control_plane import (
    ControlPlaneDeviceRegistration,
    EntraJwtTokenValidator,
    EntraJwtValidatorSettings,
    TokenValidationError,
    WindowsControlPlaneSettings,
)
from microsoft_agent.windows_control_plane_service import (
    WindowsControlPlaneServiceConfig,
    build_control_plane_service,
    load_control_plane_service_config,
    main,
    validate_control_plane_service_config,
)

TENANT_ID = UUID("11111111-1111-1111-1111-111111111111")
AUDIENCE = "00000000-0000-4000-8000-000000000010"


def _signing_material() -> tuple[Any, dict[str, Any]]:
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    public_jwk = json.loads(jwt.algorithms.RSAAlgorithm.to_jwk(key.public_key()))
    public_jwk.update({"kid": "test-signing-key", "use": "sig", "alg": "RS256"})
    return key, {"keys": [public_jwk]}


def _jwks() -> dict[str, Any]:
    return _signing_material()[1]


def _service_config(tmp_path: Path) -> WindowsControlPlaneServiceConfig:
    identity = DeviceIdentity(
        device_id="disabled-laptop",
        tenant_id=TENANT_ID,
        entra_device_id=UUID("22222222-2222-2222-2222-222222222222"),
        certificate_thumbprint="A" * 40,
    )
    device = CompanionDevice(
        identity=identity,
        display_name="Disabled test laptop",
        enabled=False,
    )
    control_plane = WindowsControlPlaneSettings(
        token_audience=AUDIENCE,
        devices={identity.device_id: ControlPlaneDeviceRegistration(device=device)},
        require_device_mtls=False,
    )
    token_validation = EntraJwtValidatorSettings(
        tenant_id=TENANT_ID,
        audience=AUDIENCE,
        jwks=_jwks(),
    )
    return WindowsControlPlaneServiceConfig(
        database_path=tmp_path / "state" / "queue.db",
        control_plane=control_plane,
        token_validation=token_validation,
    )


def test_control_plane_service_config_load_validate_and_build(
    tmp_path: Path, capsys: Any
) -> None:
    config = _service_config(tmp_path)
    raw = config.model_dump(mode="json")
    raw["database_path"] = "state/queue.db"
    config_path = tmp_path / "control-plane.json"
    config_path.write_text(json.dumps(raw), encoding="utf-8")

    loaded = load_control_plane_service_config(config_path)
    validate_control_plane_service_config(loaded)
    exit_code = main(["--config", os.fspath(config_path), "--validate-config"])
    app = build_control_plane_service(loaded)

    assert loaded.database_path == tmp_path / "state" / "queue.db"
    assert exit_code == 0
    assert "configuration is valid" in capsys.readouterr().out
    assert app.title == "Microsoft Agent Windows Companion Control Plane"
    assert loaded.database_path.is_file()


def test_control_plane_service_rejects_audience_and_unsafe_bind(
    tmp_path: Path,
) -> None:
    config = _service_config(tmp_path)
    with pytest.raises(ValidationError, match="audiences must match"):
        WindowsControlPlaneServiceConfig(
            database_path=tmp_path / "queue.db",
            control_plane=config.control_plane,
            token_validation=config.token_validation.model_copy(
                update={"audience": "api://another-api"}
            ),
        )
    with pytest.raises(ValidationError, match="Non-loopback"):
        WindowsControlPlaneServiceConfig(
            bind_host="0.0.0.0",
            database_path=tmp_path / "queue.db",
            control_plane=config.control_plane,
            token_validation=config.token_validation,
        )


def test_control_plane_service_rejects_invalid_signing_key(tmp_path: Path) -> None:
    config = _service_config(tmp_path)
    invalid = config.model_copy(
        update={
            "token_validation": config.token_validation.model_copy(
                update={
                    "jwks": {
                        "keys": [
                            {
                                "kid": "invalid",
                                "kty": "RSA",
                                "use": "sig",
                                "n": "not-a-key",
                                "e": "AQAB",
                            }
                        ]
                    }
                }
            )
        }
    )

    with pytest.raises(ValueError, match="invalid RSA key"):
        validate_control_plane_service_config(invalid)


@pytest.mark.asyncio
async def test_entra_validator_accepts_realistic_signed_v2_access_token() -> None:
    key, jwks = _signing_material()
    settings = EntraJwtValidatorSettings(
        tenant_id=TENANT_ID,
        audience=AUDIENCE,
        jwks=jwks,
    )
    now = datetime.now(UTC)
    claims = {
        "aud": AUDIENCE,
        "iss": settings.expected_issuer,
        "tid": str(TENANT_ID),
        "oid": "device-service-principal",
        "azp": "22222222-2222-4222-8222-222222222222",
        "roles": ["WindowsCompanion.Device"],
        "ver": "2.0",
        "iat": now,
        "exp": now + timedelta(minutes=5),
    }
    token = jwt.encode(
        claims,
        key,
        algorithm="RS256",
        headers={"kid": "test-signing-key"},
    )

    principal = await EntraJwtTokenValidator(settings).validate_token(token)

    assert principal.audience == AUDIENCE
    assert principal.roles == frozenset({"WindowsCompanion.Device"})

    v1_token = jwt.encode(
        {**claims, "ver": "1.0"},
        key,
        algorithm="RS256",
        headers={"kid": "test-signing-key"},
    )
    with pytest.raises(TokenValidationError, match="not an Entra v2"):
        await EntraJwtTokenValidator(settings).validate_token(v1_token)
