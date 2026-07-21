"""Runnable service host for the authenticated Windows companion control plane."""

from __future__ import annotations

import argparse
import ipaddress
import json
import os
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from microsoft_agent.windows_control_plane import (
    EntraJwtTokenValidator,
    EntraJwtValidatorSettings,
    SQLiteCompanionStore,
    WindowsControlPlaneLimits,
    WindowsControlPlaneSettings,
    create_windows_control_plane_app,
)

try:  # Optional for library-only installs.
    import uvicorn
except ImportError:  # pragma: no cover - exercised by minimal installations
    uvicorn = None  # type: ignore[assignment]

_MAXIMUM_CONFIG_BYTES = 4 * 1024 * 1024


class WindowsControlPlaneServiceConfig(BaseModel):
    """Complete, secret-free runtime configuration for the relay service."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    schema_version: Literal[1] = 1
    bind_host: str = "127.0.0.1"
    port: int = Field(default=8443, ge=1, le=65535)
    allow_non_loopback_bind: bool = False
    database_path: Path
    control_plane: WindowsControlPlaneSettings
    token_validation: EntraJwtValidatorSettings
    limits: WindowsControlPlaneLimits = Field(default_factory=WindowsControlPlaneLimits)
    log_level: Literal["critical", "error", "warning", "info"] = "info"
    access_log: bool = False

    @field_validator("bind_host")
    @classmethod
    def validate_bind_host(cls, value: str) -> str:
        try:
            return ipaddress.ip_address(value.strip()).compressed
        except ValueError as exc:
            raise ValueError("bind_host must be an explicit IP address") from exc

    @model_validator(mode="after")
    def validate_identity_boundary(self) -> WindowsControlPlaneServiceConfig:
        address = ipaddress.ip_address(self.bind_host)
        if not address.is_loopback and not self.allow_non_loopback_bind:
            raise ValueError(
                "Non-loopback control-plane binding requires "
                "allow_non_loopback_bind=true"
            )
        if self.control_plane.token_audience != self.token_validation.audience:
            raise ValueError(
                "Control-plane and token-validation audiences must match exactly"
            )
        mismatched_tenants = [
            device_id
            for device_id, registration in self.control_plane.devices.items()
            if registration.device.identity.tenant_id != self.token_validation.tenant_id
        ]
        if mismatched_tenants:
            raise ValueError(
                "Every registered device must belong to the token-validation tenant"
            )
        if str(self.database_path) == ":memory:":
            raise ValueError("The control-plane database must be durable")
        return self


def load_control_plane_service_config(
    path: str | os.PathLike[str],
) -> WindowsControlPlaneServiceConfig:
    """Load a bounded, non-symlink JSON configuration from disk."""

    config_path = Path(path)
    if config_path.is_symlink():
        raise ValueError("Control-plane config cannot be a symlink")
    if not config_path.is_file():
        raise ValueError("Control-plane config file does not exist")
    with config_path.open("rb") as stream:
        data = stream.read(_MAXIMUM_CONFIG_BYTES + 1)
    if len(data) > _MAXIMUM_CONFIG_BYTES:
        raise ValueError("Control-plane config exceeds 4 MiB")
    try:
        raw = json.loads(data)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("Control-plane config is not valid UTF-8 JSON") from exc
    if not isinstance(raw, dict):
        raise ValueError("Control-plane config must be a JSON object")
    database_value = raw.get("database_path")
    if isinstance(database_value, str):
        database_path = Path(database_value).expanduser()
        if not database_path.is_absolute():
            database_path = config_path.resolve().parent / database_path
        raw["database_path"] = os.fspath(database_path)
    return WindowsControlPlaneServiceConfig.model_validate(raw)


def validate_control_plane_service_config(
    config: WindowsControlPlaneServiceConfig,
) -> None:
    """Validate signing keys and local storage without starting a listener."""

    if config.database_path.is_symlink() or config.database_path.is_dir():
        raise ValueError("Control-plane database path must be a non-symlink file")
    parent = config.database_path.parent
    if parent.exists() and (parent.is_symlink() or not parent.is_dir()):
        raise ValueError("Control-plane database parent must be a real directory")
    EntraJwtTokenValidator(config.token_validation)


def build_control_plane_service(config: WindowsControlPlaneServiceConfig) -> Any:
    """Build the production ASGI application and durable store."""

    validate_control_plane_service_config(config)
    config.database_path.parent.mkdir(parents=True, exist_ok=True)
    store = SQLiteCompanionStore(config.database_path, config.limits)
    validator = EntraJwtTokenValidator(config.token_validation)
    return create_windows_control_plane_app(config.control_plane, store, validator)


def run_control_plane_service(config: WindowsControlPlaneServiceConfig) -> None:
    """Run one ASGI worker behind a trusted TLS/mTLS reverse proxy."""

    if uvicorn is None:
        raise ImportError("Install the control-plane extra to run the relay service")
    app = build_control_plane_service(config)
    uvicorn.run(
        app,
        host=config.bind_host,
        port=config.port,
        log_level=config.log_level,
        access_log=config.access_log,
        proxy_headers=False,
        server_header=False,
        workers=1,
    )


def main(argv: Sequence[str] | None = None) -> int:
    """Validate configuration or run the Windows control-plane service."""

    parser = argparse.ArgumentParser(
        description="Run the Microsoft Agent Windows companion control plane"
    )
    parser.add_argument(
        "--config", required=True, help="Path to the protected JSON configuration"
    )
    parser.add_argument(
        "--validate-config",
        action="store_true",
        help="Validate configuration and pinned Entra signing keys, then exit",
    )
    arguments = parser.parse_args(argv)
    try:
        config = load_control_plane_service_config(arguments.config)
        if arguments.validate_config:
            validate_control_plane_service_config(config)
            print("Windows companion control-plane configuration is valid.")
            return 0
        run_control_plane_service(config)
        return 0
    except (ImportError, OSError, RuntimeError, ValueError) as exc:
        print(f"Windows control plane failed: {type(exc).__name__}", file=sys.stderr)
        return 2


if __name__ == "__main__":  # pragma: no cover - exercised by CLI smoke tests
    raise SystemExit(main())


__all__ = [
    "WindowsControlPlaneServiceConfig",
    "build_control_plane_service",
    "load_control_plane_service_config",
    "main",
    "run_control_plane_service",
    "validate_control_plane_service_config",
]
