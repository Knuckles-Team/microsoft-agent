"""Validated runtime configuration for Microsoft Agent integrations.

The project talks to more than one Microsoft resource and supports both user
delegated and workload identities.  Keeping those values in one model prevents
the scope drift that previously existed between the MCP server, Graph wrapper,
and authentication adapter.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from enum import StrEnum
from functools import lru_cache
from pathlib import Path
from urllib.parse import urlparse

from agent_utilities.core.config import setting
from agent_utilities.security.cli_secrets import resolve_runtime_secret_reference
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    SecretStr,
    field_validator,
    model_validator,
)

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"
DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com"


class AuthenticationMode(StrEnum):
    """Supported Microsoft identity token acquisition modes."""

    DELEGATED = "delegated"
    APPLICATION = "application"
    ON_BEHALF_OF = "on_behalf_of"
    EXTERNAL_TOKEN = "external_token"
    MANAGED_IDENTITY = "managed_identity"
    WORKLOAD_IDENTITY = "workload_identity"


class LoginMethod(StrEnum):
    """Interactive login method used for delegated authentication."""

    AUTO = "auto"
    BROKER = "broker"
    BROWSER = "browser"
    DEVICE_CODE = "device_code"


class PermissionProfile(StrEnum):
    """Named least-privilege Graph permission bundles."""

    READ_ONLY = "read_only"
    PRODUCTIVITY = "productivity"
    COLLABORATION = "collaboration"
    DEVICE_READ = "device_read"
    DEVICE_ADMIN = "device_admin"
    TENANT_ADMIN = "tenant_admin"


PROFILE_SCOPES: dict[PermissionProfile, tuple[str, ...]] = {
    PermissionProfile.READ_ONLY: (
        "User.Read",
        "Mail.Read",
        "Calendars.Read",
        "Files.Read",
        "Sites.Selected",
        "Chat.Read",
        "ChannelMessage.Read.All",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "Tasks.Read",
        "Notes.Read",
        "Contacts.Read",
    ),
    PermissionProfile.PRODUCTIVITY: (
        "User.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.ReadWrite",
        "Files.ReadWrite",
        "Sites.Selected",
        "Tasks.ReadWrite",
        "Notes.ReadWrite",
        "Contacts.ReadWrite",
    ),
    PermissionProfile.COLLABORATION: (
        "User.Read",
        "Chat.ReadWrite",
        "ChatMessage.Send",
        "ChannelMessage.Read.All",
        "ChannelMessage.Send",
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "TeamMember.Read.All",
        "Group.Read.All",
        "OnlineMeetings.ReadWrite",
        "Presence.Read.All",
    ),
    PermissionProfile.DEVICE_READ: (
        "Device.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementConfiguration.Read.All",
    ),
    PermissionProfile.DEVICE_ADMIN: (
        "Device.ReadWrite.All",
        "DeviceManagementManagedDevices.ReadWrite.All",
        "DeviceManagementManagedDevices.PrivilegedOperations.All",
        "DeviceManagementConfiguration.ReadWrite.All",
    ),
    PermissionProfile.TENANT_ADMIN: (
        "Directory.ReadWrite.All",
        "Group.ReadWrite.All",
        "Application.ReadWrite.All",
        "Policy.ReadWrite.ConditionalAccess",
        "RoleManagement.ReadWrite.Directory",
        "AuditLog.Read.All",
        "Reports.Read.All",
        "SecurityIncident.ReadWrite.All",
    ),
}


DEFAULT_TOOL_GROUPS = (
    "misc",
    "auth",
    "meta",
    "mail",
    "files",
    "calendar",
    "notes",
    "tasks",
    "contacts",
    "user",
    "chat",
    "teams",
    "sites",
    "search",
    "groups",
    "communications",
    "documents",
    "power_platform",
    "windows",
    "intune",
)


def _split_values(value: str | Iterable[str] | None) -> tuple[str, ...]:
    if value is None:
        return ()
    if isinstance(value, str):
        values = re.split(r"[\s,]+", value.strip()) if value.strip() else []
    else:
        values = [str(item).strip() for item in value]
    return tuple(item for item in values if item)


def _configured_value(
    source: Mapping[str, str] | None, name: str, default: str | None = None
) -> str | None:
    if source is None:
        return setting(name, default)
    return source.get(name, default)


def _configured_bool(
    source: Mapping[str, str] | None, name: str, default: bool = False
) -> bool:
    raw = _configured_value(source, name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _resolved_secret(reference: str | None) -> SecretStr | None:
    if not reference:
        return None
    return SecretStr(resolve_runtime_secret_reference(reference))


class MicrosoftSettings(BaseModel):
    """Microsoft identity, Graph permission, and MCP safety configuration."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    tenant_id: str | None = None
    client_id: str | None = None
    authority_host: str = DEFAULT_AUTHORITY_HOST
    authentication_mode: AuthenticationMode = AuthenticationMode.DELEGATED
    login_method: LoginMethod = LoginMethod.AUTO
    permission_profiles: tuple[PermissionProfile, ...] = (
        PermissionProfile.PRODUCTIVITY,
        PermissionProfile.COLLABORATION,
    )
    custom_graph_scopes: tuple[str, ...] = ()
    client_secret: SecretStr | None = Field(default=None, repr=False)
    client_certificate_path: Path | None = None
    client_certificate_thumbprint: str | None = None
    client_certificate_password: SecretStr | None = Field(default=None, repr=False)
    managed_identity_client_id: str | None = None
    workload_identity_token_file: Path | None = None
    enable_broker: bool = True
    allow_device_code: bool = False
    require_secure_cache: bool = True
    enabled_tool_groups: tuple[str, ...] = DEFAULT_TOOL_GROUPS
    allow_write_tools: bool = False
    allow_destructive_tools: bool = False
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    graph_tls_profile: str | None = None
    graph_tls_profile_ref: str | None = None
    office_addin_origins: tuple[str, ...] = ()
    ingestion_pseudonymization_key: SecretStr | None = Field(default=None, repr=False)

    @field_validator("authority_host")
    @classmethod
    def _validate_authority_host(cls, value: str) -> str:
        origin = value.strip().rstrip("/")
        parsed = urlparse(origin)
        if (
            parsed.scheme != "https"
            or not parsed.hostname
            or parsed.username
            or parsed.password
            or parsed.path
            or parsed.params
            or parsed.query
            or parsed.fragment
        ):
            raise ValueError("Microsoft authority must be an exact HTTPS origin")
        return origin

    @field_validator("graph_base_url")
    @classmethod
    def _validate_graph_base_url(cls, value: str) -> str:
        endpoint = value.strip().rstrip("/")
        parsed = urlparse(endpoint)
        if (
            parsed.scheme != "https"
            or not parsed.hostname
            or parsed.username
            or parsed.password
            or parsed.params
            or parsed.query
            or parsed.fragment
        ):
            raise ValueError("Microsoft Graph endpoint must be an HTTPS URL")
        return endpoint

    @field_validator("graph_tls_profile")
    @classmethod
    def _validate_graph_tls_profile(cls, value: str | None) -> str | None:
        if value is None:
            return None
        selected = value.strip()
        if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,127}", selected) is None:
            raise ValueError("Microsoft Graph TLS profile name is invalid")
        return selected

    @model_validator(mode="after")
    def _validate_graph_tls_selector(self) -> MicrosoftSettings:
        if self.graph_tls_profile and self.graph_tls_profile_ref:
            raise ValueError("Microsoft Graph TLS profile selectors are ambiguous")
        return self

    @field_validator("enabled_tool_groups")
    @classmethod
    def _normalize_groups(cls, values: tuple[str, ...]) -> tuple[str, ...]:
        normalized = tuple(dict.fromkeys(value.lower().strip() for value in values))
        return normalized or DEFAULT_TOOL_GROUPS

    @field_validator("office_addin_origins")
    @classmethod
    def _validate_office_origins(cls, values: tuple[str, ...]) -> tuple[str, ...]:
        normalized: list[str] = []
        for value in values:
            origin = value.strip().rstrip("/")
            parsed = urlparse(origin)
            if (
                parsed.scheme != "https"
                or not parsed.netloc
                or parsed.username
                or parsed.password
                or parsed.path
                or parsed.query
                or parsed.fragment
            ):
                raise ValueError("Office add-in origins must be exact HTTPS origins")
            normalized.append(origin)
        return tuple(dict.fromkeys(normalized))

    @property
    def authority(self) -> str:
        """Return the tenant-specific authority URL."""

        tenant = self.tenant_id or "organizations"
        return f"{self.authority_host}/{tenant}"

    @property
    def delegated_scopes(self) -> tuple[str, ...]:
        """Return the de-duplicated delegated Graph scopes for the profiles."""

        scopes: list[str] = []
        for profile in self.permission_profiles:
            scopes.extend(PROFILE_SCOPES[profile])
        scopes.extend(self.custom_graph_scopes)
        return tuple(dict.fromkeys(scopes))

    @property
    def graph_scopes(self) -> tuple[str, ...]:
        """Return scopes appropriate to the configured authentication mode."""

        if self.authentication_mode in {
            AuthenticationMode.APPLICATION,
            AuthenticationMode.ON_BEHALF_OF,
            AuthenticationMode.MANAGED_IDENTITY,
            AuthenticationMode.WORKLOAD_IDENTITY,
        }:
            return (GRAPH_DEFAULT_SCOPE,)
        return self.delegated_scopes

    @property
    def is_configured(self) -> bool:
        """Whether the minimum identity coordinates have been supplied."""

        if self.authentication_mode is AuthenticationMode.MANAGED_IDENTITY:
            return True
        if self.authentication_mode is AuthenticationMode.WORKLOAD_IDENTITY:
            return bool(
                self.tenant_id and self.client_id and self.workload_identity_token_file
            )
        return bool(self.tenant_id and self.client_id)

    def require_identity(self) -> MicrosoftSettings:
        """Validate identity configuration immediately before authentication."""

        if self.authentication_mode is AuthenticationMode.MANAGED_IDENTITY:
            return self

        missing = [
            name
            for name, value in (
                ("MICROSOFT_TENANT_ID", self.tenant_id),
                ("MICROSOFT_CLIENT_ID", self.client_id),
            )
            if not value
        ]
        if missing:
            raise ValueError(
                "Microsoft authentication is not configured. Set "
                + " and ".join(missing)
                + "."
            )
        if self.authentication_mode is AuthenticationMode.WORKLOAD_IDENTITY:
            if not self.workload_identity_token_file:
                raise ValueError(
                    "Workload identity requires "
                    "MICROSOFT_WORKLOAD_IDENTITY_TOKEN_FILE or "
                    "AZURE_FEDERATED_TOKEN_FILE."
                )
            if not self.workload_identity_token_file.is_file():
                raise ValueError(
                    "The configured workload identity token file does not exist."
                )
            return self
        if self.authentication_mode in {
            AuthenticationMode.APPLICATION,
            AuthenticationMode.ON_BEHALF_OF,
        } and not (self.client_secret or self.client_certificate_path):
            raise ValueError(
                "Confidential authentication requires a managed workload "
                "identity, MICROSOFT_CLIENT_SECRET_REF, or a certificate. This "
                "runtime supports a secret or certificate configuration."
            )
        if self.client_certificate_path and not self.client_certificate_path.is_file():
            raise ValueError(
                "MICROSOFT_CLIENT_CERTIFICATE_PATH does not reference a file."
            )
        return self

    def tool_group_enabled(self, name: str) -> bool:
        """Return whether an MCP tool family should be registered."""

        groups = set(self.enabled_tool_groups)
        return "all" in groups or name.lower() in groups

    def require_ingestion_pseudonymization_key(self) -> bytes:
        """Return deployment-owned key material for zero-PII graph identifiers."""

        if self.ingestion_pseudonymization_key is None:
            raise ValueError(
                "Microsoft graph ingestion requires a deployment-owned "
                "pseudonymization key"
            )
        key = self.ingestion_pseudonymization_key.get_secret_value().encode("utf-8")
        if len(key) < 32:
            raise ValueError(
                "Microsoft graph ingestion pseudonymization key must contain at "
                "least 32 bytes"
            )
        return key

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> MicrosoftSettings:
        """Build settings from environment variables without mutating them."""

        profile_names = _split_values(
            _configured_value(
                env,
                "MICROSOFT_PERMISSION_PROFILES",
                "productivity,collaboration",
            )
        )
        profiles = tuple(PermissionProfile(name.lower()) for name in profile_names)
        groups = _split_values(_configured_value(env, "MICROSOFT_ENABLED_TOOL_GROUPS"))
        certificate_path = _configured_value(env, "MICROSOFT_CLIENT_CERTIFICATE_PATH")
        workload_token_file = _configured_value(
            env, "MICROSOFT_WORKLOAD_IDENTITY_TOKEN_FILE"
        ) or _configured_value(env, "AZURE_FEDERATED_TOKEN_FILE")
        return cls(
            tenant_id=_configured_value(env, "MICROSOFT_TENANT_ID"),
            client_id=_configured_value(env, "MICROSOFT_CLIENT_ID"),
            authority_host=_configured_value(
                env, "MICROSOFT_AUTHORITY_HOST", DEFAULT_AUTHORITY_HOST
            ),
            authentication_mode=AuthenticationMode(
                str(_configured_value(env, "MICROSOFT_AUTH_MODE", "delegated")).lower()
            ),
            login_method=LoginMethod(
                str(_configured_value(env, "MICROSOFT_LOGIN_METHOD", "auto")).lower()
            ),
            permission_profiles=profiles,
            custom_graph_scopes=_split_values(
                _configured_value(env, "MICROSOFT_GRAPH_SCOPES")
            ),
            client_secret=_resolved_secret(
                _configured_value(env, "MICROSOFT_CLIENT_SECRET_REF")
            ),
            client_certificate_path=(
                Path(certificate_path).expanduser() if certificate_path else None
            ),
            client_certificate_thumbprint=_configured_value(
                env, "MICROSOFT_CLIENT_CERTIFICATE_THUMBPRINT"
            ),
            client_certificate_password=_resolved_secret(
                _configured_value(env, "MICROSOFT_CLIENT_CERTIFICATE_PASSWORD_REF")
            ),
            managed_identity_client_id=_configured_value(
                env, "MICROSOFT_MANAGED_IDENTITY_CLIENT_ID"
            ),
            workload_identity_token_file=(
                Path(workload_token_file).expanduser() if workload_token_file else None
            ),
            enable_broker=_configured_bool(env, "MICROSOFT_ENABLE_BROKER", True),
            allow_device_code=_configured_bool(
                env, "MICROSOFT_ALLOW_DEVICE_CODE", False
            ),
            require_secure_cache=_configured_bool(
                env, "MICROSOFT_REQUIRE_SECURE_CACHE", True
            ),
            enabled_tool_groups=groups or DEFAULT_TOOL_GROUPS,
            allow_write_tools=_configured_bool(env, "MICROSOFT_ALLOW_WRITES", False),
            allow_destructive_tools=_configured_bool(
                env, "MICROSOFT_ALLOW_DESTRUCTIVE", False
            ),
            graph_base_url=_configured_value(
                env, "MICROSOFT_GRAPH_BASE_URL", "https://graph.microsoft.com/v1.0"
            ),
            graph_tls_profile=_configured_value(env, "MICROSOFT_GRAPH_TLS_PROFILE"),
            graph_tls_profile_ref=_configured_value(
                env, "MICROSOFT_GRAPH_TLS_PROFILE_REF"
            ),
            office_addin_origins=_split_values(
                _configured_value(env, "MICROSOFT_OFFICE_ADDIN_ORIGINS")
            ),
            ingestion_pseudonymization_key=_resolved_secret(
                _configured_value(env, "MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF")
            ),
        )


@lru_cache(maxsize=1)
def get_settings() -> MicrosoftSettings:
    """Return process-wide validated Microsoft settings."""

    return MicrosoftSettings.from_env()


def clear_settings_cache() -> None:
    """Clear cached settings; intended for tests and controlled reloads."""

    get_settings.cache_clear()
