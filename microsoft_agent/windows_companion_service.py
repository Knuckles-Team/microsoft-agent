"""Command-line service host for the outbound Windows companion worker.

Run the companion as a scheduled task or Windows service wrapper with::

    python -m microsoft_agent.windows_companion_service \
        --config <deployment-config>

The command line never accepts credentials.  The JSON config references an
ACL-protected PEM certificate/private-key pair; an optional encrypted-key
password is resolved only from the configured runtime secret reference.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import hmac
import json
import ntpath
import os
import signal
import ssl
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Any, Literal
from uuid import UUID

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    HttpUrl,
    field_validator,
    model_validator,
)

from agent_utilities.security.cli_secrets import resolve_runtime_secret_reference

from microsoft_agent.integration_auth import normalize_audience
from microsoft_agent.windows_companion import (
    CompanionActionKind,
    CompanionDevice,
    CompanionTokenProvider,
)
from microsoft_agent.windows_control_plane import HttpOutboundRelayTransport
from microsoft_agent.windows_runtime import (
    ClipboardAdapter,
    DesktopFlowExecutor,
    NotificationAdapter,
    OfficeAutomation,
    OutboundRelayTransport,
    OutboundRelayWorker,
    PyWin32ClipboardAdapter,
    PyWin32OfficeAutomation,
    PyWin32ServiceManager,
    WindowsActionExecutor,
    WindowsRuntimeLimits,
    WindowsServiceManager,
    WindowsToastNotificationAdapter,
)

try:  # Core authentication dependency, guarded for library-only installs.
    import msal
except ImportError:  # pragma: no cover - declared project dependency
    msal = None  # type: ignore[assignment]


class CompanionCertificateSettings(BaseModel):
    """Certificate credential and optional outbound mTLS file references."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    thumbprint: str = Field(pattern=r"^[A-Fa-f0-9]{40}$")
    client_certificate_path: Path
    client_private_key_path: Path
    private_key_password_ref: str | None = None
    ca_bundle_path: Path | None = None

    @field_validator("thumbprint", mode="before")
    @classmethod
    def normalize_thumbprint(cls, value: Any) -> Any:
        if isinstance(value, str):
            return value.replace(":", "").replace(" ", "").upper()
        return value


class NativeAdapterSettings(BaseModel):
    """Native integrations explicitly enabled for this device."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    office: bool = False
    services: bool = False
    clipboard: bool = False
    notifications: bool = False


class WindowsCompanionServiceConfig(BaseModel):
    """Complete fail-closed configuration for one outbound device worker."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    schema_version: Literal[1] = 1
    control_plane_url: HttpUrl
    token_audience: str = Field(min_length=1, max_length=512)
    tenant_id: UUID
    client_id: UUID
    device: CompanionDevice
    file_root_bindings: dict[str, Path] = Field(default_factory=dict)
    certificate: CompanionCertificateSettings
    adapters: NativeAdapterSettings = Field(default_factory=NativeAdapterSettings)
    runtime_limits: WindowsRuntimeLimits = Field(default_factory=WindowsRuntimeLimits)
    relay_timeout_seconds: float = Field(default=35, gt=0, le=300)
    companion_version: str = Field(
        default="microsoft-agent", min_length=1, max_length=128
    )

    @field_validator("control_plane_url")
    @classmethod
    def validate_control_plane_url(cls, value: HttpUrl) -> HttpUrl:
        if value.scheme != "https":
            raise ValueError("Control-plane URL must use HTTPS")
        if value.username or value.password:
            raise ValueError("Control-plane URL cannot contain credentials")
        if value.path not in {"", "/"} or value.query or value.fragment:
            raise ValueError("Control-plane URL must be an origin")
        return value

    @field_validator("token_audience")
    @classmethod
    def validate_token_audience(cls, value: str) -> str:
        return normalize_audience(value)

    @model_validator(mode="after")
    def validate_identity_and_roots(self) -> WindowsCompanionServiceConfig:
        if self.device.identity.tenant_id != self.tenant_id:
            raise ValueError("Configured tenant does not match device identity")
        if self.device.identity.certificate_thumbprint != self.certificate.thumbprint:
            raise ValueError("Certificate thumbprint does not match device identity")
        allowed = {
            ntpath.normcase(ntpath.normpath(item))
            for item in self.device.allowed_file_roots
        }
        supplied = {
            ntpath.normcase(ntpath.normpath(item)) for item in self.file_root_bindings
        }
        if supplied != allowed:
            raise ValueError(
                "file_root_bindings must bind every and only allowlisted file root"
            )
        return self


class MsalCertificateTokenProvider:
    """MSAL confidential-client token provider using a certificate credential."""

    def __init__(
        self,
        tenant_id: UUID,
        client_id: UUID,
        certificate: CompanionCertificateSettings,
    ) -> None:
        if msal is None:
            raise ImportError("Install msal to authenticate the Windows companion")
        private_key = _read_bounded_private_file(
            certificate.client_private_key_path, 1_048_576
        ).decode("utf-8")
        credential: dict[str, str] = {
            "thumbprint": certificate.thumbprint,
            "private_key": private_key,
        }
        password = _private_key_password(certificate)
        if password:
            credential["passphrase"] = password
        self._client_id = str(client_id)
        self._authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._credential = credential
        self._application: Any | None = None

    async def get_token(self, audience: str) -> str:
        scope = (
            audience
            if audience.endswith("/.default")
            else f"{audience.rstrip('/')}/.default"
        )
        result = await asyncio.to_thread(self._acquire_sync, scope)
        token = result.get("access_token") if isinstance(result, dict) else None
        if not isinstance(token, str) or not token:
            code = result.get("error") if isinstance(result, dict) else None
            safe_code = code if isinstance(code, str) else "token_acquisition_failed"
            raise RuntimeError(f"MSAL certificate authentication failed: {safe_code}")
        return token

    def _acquire_sync(self, scope: str) -> dict[str, Any]:
        assert msal is not None
        if self._application is None:
            self._application = msal.ConfidentialClientApplication(
                client_id=self._client_id,
                authority=self._authority,
                client_credential=self._credential,
            )
        result = self._application.acquire_token_for_client(scopes=[scope])
        return result if isinstance(result, dict) else {}


def load_service_config(path: str | os.PathLike[str]) -> WindowsCompanionServiceConfig:
    """Load a bounded, non-symlink JSON service configuration."""

    config_path = Path(path)
    if config_path.is_symlink():
        raise ValueError("Companion config cannot be a symlink")
    if not config_path.is_file():
        raise ValueError("Companion config file does not exist")
    data = _read_bounded_private_file(config_path, 1_048_576)
    try:
        return WindowsCompanionServiceConfig.model_validate_json(data)
    except (ValueError, json.JSONDecodeError) as exc:
        raise ValueError("Companion config is invalid") from exc


def validate_local_service_config(config: WindowsCompanionServiceConfig) -> None:
    """Verify referenced credentials, roots, and required native adapters."""

    if not config.device.enabled:
        return
    certificate = config.certificate
    for path, label in (
        (certificate.client_certificate_path, "client certificate"),
        (certificate.client_private_key_path, "client private key"),
    ):
        _require_secure_regular_file(path, label)
    if certificate.ca_bundle_path is not None:
        _require_secure_regular_file(certificate.ca_bundle_path, "CA bundle")
    certificate_pem = _read_bounded_private_file(
        certificate.client_certificate_path, 1_048_576
    ).decode("ascii")
    try:
        certificate_der = ssl.PEM_cert_to_DER_cert(certificate_pem)
    except ValueError as exc:
        raise ValueError("Client certificate is not valid PEM") from exc
    actual_thumbprint = hashlib.sha1(certificate_der).hexdigest().upper()  # noqa: S324
    if not hmac.compare_digest(actual_thumbprint, certificate.thumbprint):
        raise ValueError("Client certificate fingerprint does not match configuration")
    password = _private_key_password(certificate)
    context = ssl.create_default_context(
        cafile=str(certificate.ca_bundle_path) if certificate.ca_bundle_path else None
    )
    try:
        context.load_cert_chain(
            certificate.client_certificate_path,
            certificate.client_private_key_path,
            password=password,
        )
    except (OSError, ssl.SSLError) as exc:
        raise ValueError(
            "Client certificate and private key could not be loaded"
        ) from exc
    try:
        from cryptography.hazmat.primitives import serialization

        serialization.load_pem_private_key(
            _read_bounded_private_file(certificate.client_private_key_path, 1_048_576),
            password=password.encode("utf-8") if password else None,
        )
    except ImportError as exc:
        raise ImportError(
            "Install cryptography to validate the certificate private key"
        ) from exc
    except (TypeError, ValueError) as exc:
        raise ValueError("Client private key is not valid PEM") from exc
    for physical_root in config.file_root_bindings.values():
        if physical_root.is_symlink() or not physical_root.is_dir():
            raise ValueError("Every bound file root must be a real local directory")
    _validate_required_adapters(config, desktop_flows=None, validate_imports=True)


def build_worker(
    config: WindowsCompanionServiceConfig,
    *,
    token_provider: CompanionTokenProvider | None = None,
    relay_transport: OutboundRelayTransport | None = None,
    office: OfficeAutomation | None = None,
    services: WindowsServiceManager | None = None,
    clipboard: ClipboardAdapter | None = None,
    notifications: NotificationAdapter | None = None,
    desktop_flows: DesktopFlowExecutor | None = None,
) -> OutboundRelayWorker:
    """Build the authenticated worker, failing if an allowed adapter is absent."""

    if not config.device.enabled:
        raise ValueError("Companion device is disabled")
    _validate_required_adapters(
        config, desktop_flows=desktop_flows, validate_imports=False
    )
    allowed = config.device.allowed_actions
    if _needs_office(allowed) and office is None:
        office = PyWin32OfficeAutomation()
    if _needs_services(allowed) and services is None:
        services = PyWin32ServiceManager()
    if _needs_clipboard(allowed) and clipboard is None:
        clipboard = PyWin32ClipboardAdapter()
    if CompanionActionKind.NOTIFICATION_SHOW in allowed and notifications is None:
        notifications = WindowsToastNotificationAdapter()

    if relay_transport is None:
        certificate = config.certificate
        _require_secure_regular_file(
            certificate.client_certificate_path, "client certificate"
        )
        _require_secure_regular_file(
            certificate.client_private_key_path, "client private key"
        )
        if token_provider is None:
            token_provider = MsalCertificateTokenProvider(
                config.tenant_id, config.client_id, certificate
            )
        password = _private_key_password(certificate)
        relay_transport = HttpOutboundRelayTransport(
            str(config.control_plane_url).rstrip("/"),
            config.token_audience,
            token_provider,
            timeout_seconds=config.relay_timeout_seconds,
            companion_version=config.companion_version,
            capabilities=frozenset(allowed),
            client_certificate_path=certificate.client_certificate_path,
            client_private_key_path=certificate.client_private_key_path,
            private_key_password=password,
            ca_bundle_path=certificate.ca_bundle_path,
        )

    executor = WindowsActionExecutor(
        config.device,
        file_root_bindings=config.file_root_bindings,
        limits=config.runtime_limits,
        office=office,
        services=services,
        clipboard=clipboard,
        notifications=notifications,
        desktop_flows=desktop_flows,
    )
    return OutboundRelayWorker(relay_transport, executor, limits=config.runtime_limits)


def _private_key_password(certificate: CompanionCertificateSettings) -> str | None:
    if not certificate.private_key_password_ref:
        return None
    return resolve_runtime_secret_reference(certificate.private_key_password_ref)


async def run_service(config: WindowsCompanionServiceConfig) -> None:
    """Run the configured outbound worker until SIGINT or SIGTERM."""

    worker = build_worker(config)
    stop_event = asyncio.Event()
    loop = asyncio.get_running_loop()

    def stop(*_args: Any) -> None:
        loop.call_soon_threadsafe(stop_event.set)

    previous: dict[signal.Signals, Any] = {}
    watched_signals = (signal.Signals.SIGINT, signal.Signals.SIGTERM)
    for signum in watched_signals:
        try:
            previous[signum] = signal.signal(signum, stop)
        except (ValueError, OSError):
            pass
    try:
        await worker.run(stop_event)
    finally:
        for registered_signal, handler in previous.items():
            try:
                signal.signal(registered_signal, handler)
            except (ValueError, OSError):
                pass


def main(argv: Sequence[str] | None = None) -> int:
    """Validate configuration or run the Windows companion service."""

    parser = argparse.ArgumentParser(
        description="Run the authenticated outbound Microsoft Windows companion"
    )
    parser.add_argument(
        "--config", required=True, help="Path to the ACL-protected JSON config"
    )
    parser.add_argument(
        "--validate-config",
        action="store_true",
        help="Validate config, credentials, roots, and adapters, then exit",
    )
    arguments = parser.parse_args(argv)
    try:
        config = load_service_config(arguments.config)
        if arguments.validate_config:
            validate_local_service_config(config)
            if config.device.enabled:
                print("Windows companion configuration is valid and enabled.")
            else:
                print(
                    "Windows companion configuration schema is valid; device is disabled."
                )
            return 0
        asyncio.run(run_service(config))
        return 0
    except (ImportError, OSError, RuntimeError, ValueError) as exc:
        print(f"Windows companion failed: {type(exc).__name__}", file=sys.stderr)
        return 2


def _validate_required_adapters(
    config: WindowsCompanionServiceConfig,
    *,
    desktop_flows: DesktopFlowExecutor | None,
    validate_imports: bool,
) -> None:
    allowed = config.device.allowed_actions
    if _needs_office(allowed) and not config.adapters.office:
        raise ValueError("Office actions require the Office adapter to be enabled")
    if _needs_services(allowed) and not config.adapters.services:
        raise ValueError("Service actions require the services adapter to be enabled")
    if _needs_clipboard(allowed) and not config.adapters.clipboard:
        raise ValueError(
            "Clipboard actions require the clipboard adapter to be enabled"
        )
    if (
        CompanionActionKind.NOTIFICATION_SHOW in allowed
        and not config.adapters.notifications
    ):
        raise ValueError(
            "Notification actions require the notification adapter to be enabled"
        )
    if (
        CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN in allowed
        and desktop_flows is None
    ):
        raise ValueError(
            "Power Automate Desktop is allowlisted but no typed executor was injected"
        )
    if validate_imports:
        if _needs_office(allowed):
            PyWin32OfficeAutomation()
        if _needs_services(allowed):
            PyWin32ServiceManager()
        if _needs_clipboard(allowed):
            PyWin32ClipboardAdapter()
        if CompanionActionKind.NOTIFICATION_SHOW in allowed:
            WindowsToastNotificationAdapter()


def _needs_office(actions: frozenset[CompanionActionKind]) -> bool:
    return bool(
        actions
        & {
            CompanionActionKind.OFFICE_OPEN_DOCUMENT,
            CompanionActionKind.OFFICE_EXPORT_PDF,
        }
    )


def _needs_services(actions: frozenset[CompanionActionKind]) -> bool:
    return bool(
        actions
        & {
            CompanionActionKind.WINDOWS_SERVICE_STATUS,
            CompanionActionKind.WINDOWS_SERVICE_START,
            CompanionActionKind.WINDOWS_SERVICE_STOP,
        }
    )


def _needs_clipboard(actions: frozenset[CompanionActionKind]) -> bool:
    return bool(
        actions
        & {
            CompanionActionKind.CLIPBOARD_READ_TEXT,
            CompanionActionKind.CLIPBOARD_WRITE_TEXT,
        }
    )


def _read_bounded_private_file(path: Path, maximum_bytes: int) -> bytes:
    if path.is_symlink():
        raise ValueError(f"Refusing symlink file: {path.name}")
    with path.open("rb") as stream:
        data = stream.read(maximum_bytes + 1)
    if len(data) > maximum_bytes:
        raise ValueError(f"File exceeds its size limit: {path.name}")
    return data


def _require_secure_regular_file(path: Path, label: str) -> None:
    if path.is_symlink() or not path.is_file():
        raise ValueError(f"Configured {label} must be a regular non-symlink file")
    if os.name != "nt" and label == "client private key":
        mode = path.stat().st_mode & 0o777
        if mode & 0o077:
            raise ValueError(
                "Client private key permissions must deny group/other access"
            )


if __name__ == "__main__":  # pragma: no cover - exercised by deployment smoke tests
    raise SystemExit(main())


__all__ = [
    "CompanionCertificateSettings",
    "MsalCertificateTokenProvider",
    "NativeAdapterSettings",
    "WindowsCompanionServiceConfig",
    "build_worker",
    "load_service_config",
    "main",
    "run_service",
    "validate_local_service_config",
]
