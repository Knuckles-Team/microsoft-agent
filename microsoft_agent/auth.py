"""Microsoft identity authentication for delegated and workload scenarios."""

from __future__ import annotations

import atexit
import base64
import json
import logging
import sys
import time
from collections.abc import AsyncIterator, Callable, Sequence
from pathlib import Path
from typing import Any

import keyring
import msal
from agent_utilities.core.config import setting
from agent_utilities.core.exceptions import AuthError, UnauthorizedError
from keyring.errors import KeyringError

from microsoft_agent.settings import (
    AuthenticationMode,
    LoginMethod,
    MicrosoftSettings,
    get_settings,
)

logger = logging.getLogger(__name__)

SERVICE_NAME = "microsoft-agent-mcp"
TOKEN_CACHE_ACCOUNT = "msal_token_cache"
SELECTED_ACCOUNT_KEY = "selected_account"


class AuthenticationRequiredError(ValueError):
    """Raised when an operation requires a Microsoft access token."""


def _safe_token_error(result: dict[str, Any] | None) -> str:
    """Return a stable diagnostic without provider text or correlation values."""

    return "Authentication failed: Microsoft identity rejected the request"


def _jwt_expiry(token: str) -> int:
    """Read an access token expiry without treating the token as verified."""

    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        decoded = json.loads(base64.urlsafe_b64decode(payload).decode("utf-8"))
        return int(decoded.get("exp", int(time.time()) + 300))
    except (IndexError, ValueError, TypeError, json.JSONDecodeError):
        return int(time.time()) + 300


def _jwt_claims(token: str) -> dict[str, Any] | None:
    """Decode unverified JWT claims for defensive routing checks only."""

    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload).decode("utf-8"))
        return claims if isinstance(claims, dict) else None
    except (IndexError, ValueError, TypeError, json.JSONDecodeError):
        return None


def _default_scope_audiences(scopes: Sequence[str]) -> frozenset[str]:
    """Extract explicit resource audiences from OAuth ``/.default`` scopes."""

    suffix = "/.default"
    return frozenset(
        scope[: -len(suffix)].rstrip("/")
        for scope in scopes
        if scope.endswith(suffix) and len(scope) > len(suffix)
    )


def _get_incoming_user_token() -> str | None:
    """Read the request token captured by agent-utilities middleware."""

    try:
        from agent_utilities.mcp.delegated_auth import get_user_token

        return get_user_token()
    except (ImportError, AttributeError):
        return None


class AuthManager:
    """Acquire and cache Microsoft tokens for one client and authority.

    Callers should use :func:`build_auth_manager` so authentication mode and
    credential policy come from
    :class:`~microsoft_agent.settings.MicrosoftSettings`.
    """

    def __init__(
        self,
        client_id: str,
        authority: str,
        scopes: list[str] | tuple[str, ...],
        *,
        mode: AuthenticationMode | str = AuthenticationMode.DELEGATED,
        client_credential: str | dict[str, Any] | None = None,
        enable_broker: bool = False,
        allow_device_code: bool = False,
        require_secure_cache: bool = True,
        external_token_provider: Callable[[], str | None] | None = None,
        azure_credential: Any | None = None,
        expected_tenant_id: str | None = None,
        allowed_external_audiences: Sequence[str] = (),
        graph_base_url: str = "https://graph.microsoft.com/v1.0",
        graph_tls_profile: str | None = None,
        graph_tls_profile_ref: str | None = None,
    ):
        self.client_id = client_id
        self.authority = authority
        self.scopes = list(scopes)
        self.mode = AuthenticationMode(mode)
        self.enable_broker = enable_broker
        self.allow_device_code = allow_device_code
        self.require_secure_cache = require_secure_cache
        self.external_token_provider = external_token_provider
        self.azure_credential = azure_credential
        self.expected_tenant_id = expected_tenant_id
        self.allowed_external_audiences = frozenset(allowed_external_audiences)
        self.graph_base_url = graph_base_url
        self.graph_tls_profile = graph_tls_profile
        self.graph_tls_profile_ref = graph_tls_profile_ref
        self.token_cache = msal.SerializableTokenCache()
        self.access_token: str | None = None
        self.selected_account_id: str | None = None
        self.secure_cache_available = True

        if self.mode is AuthenticationMode.DELEGATED:
            self.load_token_cache()
        self.msal_app = self._create_msal_application(client_credential)
        if setting("TESTING") != "1":
            atexit.register(self.save_token_cache)

    def _create_msal_application(
        self, client_credential: str | dict[str, Any] | None
    ) -> Any:
        if self.mode in {
            AuthenticationMode.EXTERNAL_TOKEN,
            AuthenticationMode.MANAGED_IDENTITY,
            AuthenticationMode.WORKLOAD_IDENTITY,
        }:
            if (
                self.mode
                in {
                    AuthenticationMode.MANAGED_IDENTITY,
                    AuthenticationMode.WORKLOAD_IDENTITY,
                }
                and self.azure_credential is None
            ):
                raise ValueError(
                    "Azure workload authentication requires an Azure credential."
                )
            return None
        if self.mode in {
            AuthenticationMode.APPLICATION,
            AuthenticationMode.ON_BEHALF_OF,
        }:
            if not client_credential:
                raise ValueError(
                    "Confidential authentication requires a client credential."
                )
            return msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=client_credential,
                token_cache=self.token_cache,
            )

        kwargs: dict[str, Any] = {
            "authority": self.authority,
            "token_cache": self.token_cache,
        }
        broker_requested = self.enable_broker and sys.platform == "win32"
        if broker_requested:
            kwargs["enable_broker_on_windows"] = True
        return msal.PublicClientApplication(self.client_id, **kwargs)

    def load_token_cache(self) -> None:
        """Load the token cache from OS-backed secret storage only."""

        cache_data = None
        try:
            cache_data = keyring.get_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT)
        except (KeyringError, ImportError) as exc:
            self.secure_cache_available = False
            logger.warning(
                "Secure token storage is unavailable: error_type=%s",
                type(exc).__name__,
            )

        if cache_data:
            try:
                self.token_cache.deserialize(cache_data)
            except (ValueError, TypeError, json.JSONDecodeError):
                logger.warning("Ignoring an invalid cached Microsoft token payload")

        self.load_selected_account()

    def save_token_cache(self) -> None:
        """Persist the token cache only in OS-backed secret storage."""

        if self.token_cache.has_state_changed:
            cache_data = self.token_cache.serialize()
            try:
                keyring.set_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cache_data)
                self.secure_cache_available = True
            except (KeyringError, ImportError) as exc:
                self.secure_cache_available = False
                logger.warning(
                    "Token cache remains in memory because secure storage is "
                    "unavailable: error_type=%s",
                    type(exc).__name__,
                )

        self.save_selected_account()

    def load_selected_account(self) -> None:
        """Load the selected delegated account from secure storage."""

        try:
            data = keyring.get_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY)
        except (KeyringError, ImportError):
            self.secure_cache_available = False
            return

        if data:
            try:
                parsed = json.loads(data)
                self.selected_account_id = parsed.get("account_id")
            except (json.JSONDecodeError, TypeError):
                logger.warning("Ignoring an invalid selected-account payload")

    def save_selected_account(self) -> None:
        """Persist delegated account selection in secure storage."""

        if not self.selected_account_id:
            return
        data = json.dumps({"account_id": self.selected_account_id})
        try:
            keyring.set_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY, data)
            self.secure_cache_available = True
        except (KeyringError, ImportError) as exc:
            self.secure_cache_available = False
            logger.warning(
                "Account selection remains in memory because secure storage is "
                "unavailable: error_type=%s",
                type(exc).__name__,
            )

    def get_current_account(self) -> dict[str, Any] | None:
        """Return the selected delegated account, if one is cached."""

        if self.mode is not AuthenticationMode.DELEGATED:
            return None
        accounts = self.msal_app.get_accounts()
        if not accounts:
            return None
        if self.selected_account_id:
            for account in accounts:
                if account.get("home_account_id") == self.selected_account_id:
                    return account
            logger.warning("Selected Microsoft account is unavailable")
            return None
        return accounts[0]

    def _requested_scopes(self, scopes: Sequence[str] | None) -> list[str]:
        requested = [scope for scope in (scopes or self.scopes) if scope]
        return list(dict.fromkeys(requested))

    def get_token(self, scopes: Sequence[str] | None = None) -> str | None:
        """Acquire a cached or non-interactive token for the requested scopes."""

        details = self.get_token_details(scopes=scopes)
        if details and "access_token" in details:
            return str(details["access_token"])
        return None

    def acquire_token_for_scopes(
        self,
        scopes: Sequence[str],
        *,
        allow_interactive: bool = False,
    ) -> str | None:
        """Acquire a token for another configured resource audience.

        Delegated multi-resource consent cannot be requested in one OAuth token
        operation. Callers must opt into interactive acquisition explicitly;
        workload modes remain non-interactive.
        """

        requested = self._requested_scopes(scopes)
        token = self.get_token(requested)
        if token or not allow_interactive:
            return token
        if self.mode is not AuthenticationMode.DELEGATED:
            return None
        self._require_secure_delegated_cache()
        result = self.msal_app.acquire_token_interactive(scopes=requested)
        self._accept_interactive_result(result)
        return str(result["access_token"]) if result else None

    def get_token_details(
        self, scopes: Sequence[str] | None = None
    ) -> dict[str, Any] | None:
        """Return the full MSAL token result without prompting the user."""

        requested = self._requested_scopes(scopes)
        if self.mode is AuthenticationMode.EXTERNAL_TOKEN:
            token = (
                self.external_token_provider() if self.external_token_provider else None
            )
            if not token:
                return None
            if self.expected_tenant_id or self.allowed_external_audiences:
                claims = _jwt_claims(token)
                if claims is None:
                    logger.warning("External token is not a parseable JWT")
                    return None
                tenant = claims.get("tid")
                audience = claims.get("aud")
                audiences = (
                    {str(item) for item in audience}
                    if isinstance(audience, list)
                    else {str(audience)}
                )
                if self.expected_tenant_id and tenant != self.expected_tenant_id:
                    logger.warning("External token tenant does not match configuration")
                    return None
                if self.allowed_external_audiences and not (
                    audiences & self.allowed_external_audiences
                ):
                    logger.warning("External token audience is not Microsoft Graph")
                    return None
                requested_audiences = _default_scope_audiences(requested)
                if (
                    requested_audiences
                    and self.allowed_external_audiences
                    and not requested_audiences.issubset(
                        self.allowed_external_audiences
                    )
                ):
                    logger.warning(
                        "External token cannot be reused for another resource audience"
                    )
                    return None
            return {"access_token": token, "expires_on": _jwt_expiry(token)}

        if self.mode is AuthenticationMode.APPLICATION:
            result = self.msal_app.acquire_token_for_client(scopes=requested)
            return result if result and "access_token" in result else None

        if self.mode in {
            AuthenticationMode.MANAGED_IDENTITY,
            AuthenticationMode.WORKLOAD_IDENTITY,
        }:
            credential = self.azure_credential
            if credential is None:
                logger.warning(
                    "Azure workload credential is unavailable: mode=%s",
                    self.mode.value,
                )
                return None
            try:
                azure_token = credential.get_token(*requested)
            except Exception:
                logger.warning(
                    "Azure workload token acquisition failed: mode=%s",
                    self.mode.value,
                )
                return None
            return {
                "access_token": azure_token.token,
                "expires_on": int(azure_token.expires_on),
            }

        if self.mode is AuthenticationMode.ON_BEHALF_OF:
            assertion = (
                self.external_token_provider() if self.external_token_provider else None
            )
            if not assertion:
                return None
            result = self.msal_app.acquire_token_on_behalf_of(
                user_assertion=assertion,
                scopes=requested,
            )
            return result if result and "access_token" in result else None

        account = self.get_current_account()
        if not account:
            return None
        result = self.msal_app.acquire_token_silent(requested, account=account)
        return result if result and "access_token" in result else None

    def _accept_interactive_result(self, result: dict[str, Any] | None) -> str:
        if not result or "access_token" not in result:
            raise AuthError(_safe_token_error(result))
        self.access_token = str(result["access_token"])
        account = result.get("account")
        home_account_id = result.get("home_account_id")
        if not home_account_id and isinstance(account, dict):
            home_account_id = account.get("home_account_id")
        if home_account_id:
            self.selected_account_id = str(home_account_id)
        self.save_token_cache()
        return "Authentication successful"

    def acquire_token_by_device_code(self, callback: Callable[[str], None]) -> str:
        """Authenticate a delegated user through device authorization flow."""

        if self.mode is not AuthenticationMode.DELEGATED:
            raise ValueError("Device code login is only valid for delegated auth.")
        self._require_secure_delegated_cache()
        if not self.allow_device_code:
            raise ValueError(
                "Device code login is disabled. Set MICROSOFT_ALLOW_DEVICE_CODE=1 "
                "only when tenant policy permits this higher-risk flow."
            )
        flow = self.msal_app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise AuthError("Failed to create device flow")
        callback(str(flow.get("message", "Complete Microsoft device login.")))
        result = self.msal_app.acquire_token_by_device_flow(flow)
        return self._accept_interactive_result(result)

    def acquire_token_interactive(self) -> str:
        """Authenticate with the system browser or Windows broker."""

        if self.mode is not AuthenticationMode.DELEGATED:
            raise ValueError("Interactive login is only valid for delegated auth.")
        self._require_secure_delegated_cache()
        result = self.msal_app.acquire_token_interactive(scopes=self.scopes)
        return self._accept_interactive_result(result)

    def _require_secure_delegated_cache(self) -> None:
        """Fail before prompting when required OS secret storage is unavailable."""

        if self.require_secure_cache and not self.secure_cache_available:
            raise AuthError(
                "Secure OS token storage is unavailable. Configure the platform "
                "keyring or set MICROSOFT_REQUIRE_SECURE_CACHE=false for an "
                "explicit memory-only session."
            )

    def login(
        self,
        *,
        method: LoginMethod | str = LoginMethod.AUTO,
        callback: Callable[[str], None] = print,
    ) -> str:
        """Authenticate using the configured delegated login method."""

        if self.mode is not AuthenticationMode.DELEGATED:
            if self.get_token():
                return "Authentication successful"
            raise AuthError("Configured Microsoft workload identity is unavailable")

        selected = LoginMethod(method)
        if selected is LoginMethod.AUTO:
            if sys.platform == "win32" and self.enable_broker:
                selected = LoginMethod.BROKER
            else:
                selected = LoginMethod.BROWSER
        if selected is LoginMethod.DEVICE_CODE:
            return self.acquire_token_by_device_code(callback)
        return self.acquire_token_interactive()

    def logout(self) -> None:
        """Remove cached delegated accounts and secure cache entries."""

        if self.mode is AuthenticationMode.DELEGATED:
            for account in self.msal_app.get_accounts():
                self.msal_app.remove_account(account)
        self.selected_account_id = None
        self.access_token = None
        try:
            keyring.delete_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT)
            keyring.delete_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY)
        except Exception:  # Keyring backends use several provider exceptions.
            pass

    def list_accounts(self) -> list[dict[str, Any]]:
        """List cached delegated accounts."""

        if self.mode is not AuthenticationMode.DELEGATED:
            return []
        return list(self.msal_app.get_accounts())

    def select_account(self, account_id: str) -> bool:
        """Select a delegated account by MSAL home-account ID."""

        for account in self.list_accounts():
            if account.get("home_account_id") == account_id:
                self.selected_account_id = account_id
                self.save_selected_account()
                return True
        return False

    def remove_account(self, account_id: str) -> bool:
        """Remove one delegated account from the token cache."""

        for account in self.list_accounts():
            if account.get("home_account_id") == account_id:
                self.msal_app.remove_account(account)
                if self.selected_account_id == account_id:
                    self.selected_account_id = None
                self.save_token_cache()
                return True
        return False


def _certificate_credential(settings: MicrosoftSettings) -> dict[str, Any]:
    path = Path(settings.client_certificate_path or "")
    private_key = path.read_text(encoding="utf-8")
    if not settings.client_certificate_thumbprint:
        raise ValueError(
            "MICROSOFT_CLIENT_CERTIFICATE_THUMBPRINT is required for PEM "
            "certificate authentication."
        )
    credential: dict[str, Any] = {
        "private_key": private_key,
        "thumbprint": settings.client_certificate_thumbprint,
    }
    if settings.client_certificate_password:
        credential["passphrase"] = (
            settings.client_certificate_password.get_secret_value()
        )
    return credential


def _azure_identity_credential(settings: MicrosoftSettings) -> Any | None:
    """Create an optional Azure-hosted identity credential lazily."""

    if settings.authentication_mode not in {
        AuthenticationMode.MANAGED_IDENTITY,
        AuthenticationMode.WORKLOAD_IDENTITY,
    }:
        return None
    try:
        from azure.identity import ManagedIdentityCredential, WorkloadIdentityCredential
    except ImportError as exc:
        raise ImportError(
            "Install microsoft-agent[cloud] to use Azure managed or workload identity."
        ) from exc

    if settings.authentication_mode is AuthenticationMode.MANAGED_IDENTITY:
        kwargs = (
            {"client_id": settings.managed_identity_client_id}
            if settings.managed_identity_client_id
            else {}
        )
        return ManagedIdentityCredential(**kwargs)
    return WorkloadIdentityCredential(
        tenant_id=str(settings.tenant_id),
        client_id=str(settings.client_id),
        token_file_path=str(settings.workload_identity_token_file),
    )


def build_auth_manager(settings: MicrosoftSettings | None = None) -> AuthManager:
    """Construct an authentication manager from validated runtime settings."""

    config = (settings or get_settings()).require_identity()
    credential: str | dict[str, Any] | None = None
    if config.client_certificate_path:
        credential = _certificate_credential(config)
    elif config.client_secret:
        credential = config.client_secret.get_secret_value()
    provider = (
        _get_incoming_user_token
        if config.authentication_mode
        in {AuthenticationMode.ON_BEHALF_OF, AuthenticationMode.EXTERNAL_TOKEN}
        else None
    )
    azure_credential = _azure_identity_credential(config)
    return AuthManager(
        client_id=str(config.client_id or config.managed_identity_client_id or ""),
        authority=config.authority,
        scopes=config.graph_scopes,
        mode=config.authentication_mode,
        client_credential=credential,
        enable_broker=config.enable_broker,
        allow_device_code=config.allow_device_code,
        require_secure_cache=config.require_secure_cache,
        external_token_provider=provider,
        azure_credential=azure_credential,
        expected_tenant_id=config.tenant_id,
        allowed_external_audiences=(
            "https://graph.microsoft.com",
            "00000003-0000-0000-c000-000000000000",
        ),
        graph_base_url=config.graph_base_url,
        graph_tls_profile=config.graph_tls_profile,
        graph_tls_profile_ref=config.graph_tls_profile_ref,
    )


_auth_manager: AuthManager | None = None
_auth_fingerprint: tuple[Any, ...] | None = None


def get_auth_manager(
    settings: MicrosoftSettings | None = None, *, refresh: bool = False
) -> AuthManager:
    """Return a cached authentication manager for the active configuration."""

    global _auth_fingerprint, _auth_manager
    config = settings or get_settings()
    fingerprint = (
        config.tenant_id,
        config.client_id,
        config.authentication_mode,
        config.graph_scopes,
        config.client_certificate_path,
        bool(config.client_secret),
        config.enable_broker,
        config.allow_device_code,
        config.managed_identity_client_id,
        config.workload_identity_token_file,
        config.graph_base_url,
        config.graph_tls_profile,
        config.graph_tls_profile_ref,
    )
    if refresh or _auth_manager is None or fingerprint != _auth_fingerprint:
        _auth_manager = build_auth_manager(config)
        _auth_fingerprint = fingerprint
    return _auth_manager


def clear_auth_manager_cache() -> None:
    """Discard the process-wide manager; intended for tests and config reloads."""

    global _auth_fingerprint, _auth_manager
    _auth_manager = None
    _auth_fingerprint = None


def login(
    *,
    force: bool = False,
    method: LoginMethod | str | None = None,
    callback: Callable[[str], None] = print,
) -> str:
    """Bootstrap authentication without requiring an existing Graph client."""

    settings = get_settings().require_identity()
    manager = get_auth_manager(settings, refresh=force)
    if not force and manager.get_token():
        return "Already authenticated."
    return manager.login(method=method or settings.login_method, callback=callback)


async def get_client() -> Any:
    """Build an authenticated Microsoft Graph wrapper."""

    from microsoft_agent.api_client import MicrosoftGraphApi

    try:
        manager = get_auth_manager()
        if not manager.get_token():
            raise AuthenticationRequiredError(
                "Microsoft token is not available. Use the login tool for "
                "delegated mode or configure workload authentication."
            )
        return MicrosoftGraphApi(manager)
    except (AuthError, UnauthorizedError) as exc:
        raise RuntimeError(
            "AUTHENTICATION ERROR: Microsoft credentials were rejected."
        ) from exc


async def get_client_dependency() -> AsyncIterator[Any]:
    """Yield one request-scoped Graph client and always release its transport."""

    client = await get_client()
    try:
        yield client
    finally:
        await client.close()
