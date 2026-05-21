"""Microsoft Agent Authentication Module.

Authentication priority:
1. **OIDC Delegation** — If ``ENABLE_DELEGATION`` is active, exchanges
   the IdP-issued user token for a downstream Microsoft Graph access
   token via RFC 8693 Token Exchange.
2. **MSAL Device Code Flow** — Interactive device code flow via the
   ``AuthManager`` class using Microsoft's MSAL library.
3. **MSAL Token Cache** — Silent token acquisition from cached MSAL
   tokens (via keyring or file fallback).
4. **MCP User Token** — Direct use of the ``user_token`` from
   ``UserTokenMiddleware`` as a Microsoft Graph bearer token.

See ``docs/guides/oauth_sso.md`` in agent-utilities for full details.
"""

import atexit
import json
import logging
import os
from pathlib import Path
from typing import TYPE_CHECKING, Any

import keyring
import msal
from agent_utilities.exceptions import AuthError, UnauthorizedError
from keyring.errors import KeyringError

if TYPE_CHECKING:
    pass

logger = logging.getLogger(__name__)

SERVICE_NAME = "microsoft-agent-mcp"
TOKEN_CACHE_ACCOUNT = "msal_token_cache"  # nosec B105
SELECTED_ACCOUNT_KEY = "selected_account"
FALLBACK_DIR = Path.home() / ".microsoft-agent"
FALLBACK_PATH = FALLBACK_DIR / ".token_cache.json"
SELECTED_ACCOUNT_PATH = FALLBACK_DIR / ".selected_account.json"

FALLBACK_DIR.mkdir(parents=True, exist_ok=True)


class AuthManager:
    def __init__(self, client_id: str, authority: str, scopes: list[str]):
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.token_cache = msal.SerializableTokenCache()
        self.msal_app = msal.PublicClientApplication(
            self.client_id, authority=self.authority, token_cache=self.token_cache
        )
        self.access_token: str | None = None
        self.selected_account_id: str | None = None

        self.load_token_cache()
        atexit.register(self.save_token_cache)

    def load_token_cache(self):
        """Load token cache from keyring or file fallback."""
        cache_data = None
        try:
            cache_data = keyring.get_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT)
        except (KeyringError, ImportError) as e:
            logger.warning(f"Keyring access failed: {e}. Using file fallback.")

        if not cache_data and FALLBACK_PATH.exists():
            try:
                with open(FALLBACK_PATH) as f:
                    cache_data = f.read()
            except Exception as e:
                logger.error(f"Failed to read token cache file: {e}")

        if cache_data:
            self.token_cache.deserialize(cache_data)

        self.load_selected_account()

    def save_token_cache(self):
        """Save token cache to keyring or file fallback."""
        if self.token_cache.has_state_changed:
            cache_data = self.token_cache.serialize()
            try:
                keyring.set_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, cache_data)
            except (KeyringError, ImportError) as e:
                logger.warning(f"Keyring save failed: {e}. Using file fallback.")
                try:
                    with open(FALLBACK_PATH, "w") as f:
                        f.write(cache_data)
                except Exception as ex:
                    logger.error(f"Failed to write token cache file: {ex}")

        self.save_selected_account()

    def load_selected_account(self):
        """Load selected account ID."""
        data = None
        try:
            data = keyring.get_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY)
        except (KeyringError, ImportError):
            pass

        if not data and SELECTED_ACCOUNT_PATH.exists():
            try:
                with open(SELECTED_ACCOUNT_PATH) as f:
                    data = f.read()
            except Exception as e:
                logger.error(f"Failed to read selected account file: {e}")

        if data:
            try:
                parsed = json.loads(data)
                self.selected_account_id = parsed.get("account_id")
            except json.JSONDecodeError:
                pass

    def save_selected_account(self):
        """Save selected account ID."""
        if not self.selected_account_id:
            return

        data = json.dumps({"account_id": self.selected_account_id})
        try:
            keyring.set_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY, data)
        except (KeyringError, ImportError):
            try:
                with open(SELECTED_ACCOUNT_PATH, "w") as f:
                    f.write(data)
            except Exception as e:
                logger.error(f"Failed to write selected account file: {e}")

    def get_current_account(self) -> dict[str, Any] | None:
        accounts = self.msal_app.get_accounts()
        if not accounts:
            return None

        if self.selected_account_id:
            for acc in accounts:
                if acc.get("home_account_id") == self.selected_account_id:
                    return acc
            logger.warning(
                f"Selected account {self.selected_account_id} not found, falling back to first account."
            )

        return accounts[0]

    def get_token(self) -> str | None:
        account = self.get_current_account()
        if not account:
            return None

        result = self.msal_app.acquire_token_silent(self.scopes, account=account)
        if result and "access_token" in result:
            return result["access_token"]

        return None

    def get_token_details(
        self,
        claims: str | None = None,
        tenant_id: str | None = None,
        **kwargs: Any,
    ) -> dict[str, Any] | None:
        """Get the full token response from MSAL."""
        account = self.get_current_account()
        if not account:
            return None

        authority = (
            f"https://login.microsoftonline.com/{tenant_id}" if tenant_id else None
        )

        result = self.msal_app.acquire_token_silent(
            self.scopes,
            account=account,
            claims_challenge=claims,
            authority=authority,
            **kwargs,
        )
        if result and "access_token" in result:
            return result

        return None

    def acquire_token_by_device_code(self, callback) -> str:
        """Initiate device code flow."""
        flow = self.msal_app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise Exception("Failed to create device flow")

        callback(flow["message"])

        result = self.msal_app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            self.access_token = result["access_token"]

            if not self.selected_account_id:
                pass

            self.save_token_cache()

            if "home_account_id" in result:
                self.selected_account_id = result["home_account_id"]
                self.save_selected_account()

            return "Authentication successful"
        else:
            error = result.get("error_description", "Unknown error")
            raise Exception(f"Authentication failed: {error}")

    def logout(self):
        accounts = self.msal_app.get_accounts()
        for acc in accounts:
            self.msal_app.remove_account(acc)

        self.selected_account_id = None
        self.access_token = None

        try:
            keyring.delete_password(SERVICE_NAME, TOKEN_CACHE_ACCOUNT)
            keyring.delete_password(SERVICE_NAME, SELECTED_ACCOUNT_KEY)
        except Exception:  # nosec B110
            pass

        if FALLBACK_PATH.exists():
            FALLBACK_PATH.unlink()
        if SELECTED_ACCOUNT_PATH.exists():
            SELECTED_ACCOUNT_PATH.unlink()

    def list_accounts(self) -> list[dict[str, Any]]:
        return self.msal_app.get_accounts()

    def select_account(self, account_id: str) -> bool:
        accounts = self.msal_app.get_accounts()
        for acc in accounts:
            if acc.get("home_account_id") == account_id:
                self.selected_account_id = account_id
                self.save_selected_account()
                return True
        return False

    def remove_account(self, account_id: str) -> bool:
        accounts = self.msal_app.get_accounts()
        for acc in accounts:
            if acc.get("home_account_id") == account_id:
                self.msal_app.remove_account(acc)
                if self.selected_account_id == account_id:
                    self.selected_account_id = None
                    self.save_selected_account()
                self.save_token_cache()
                return True
        return False


async def get_client():
    from microsoft_agent.api_client import MicrosoftGraphApi

    """Create a Microsoft Graph API client with the best available auth method.

    Authentication priority:
    1. OIDC Delegation (RFC 8693 Token Exchange) — if ENABLE_DELEGATION is True
    2. MSAL cached token — silent acquisition from keyring/file cache
    3. MCP User Token — direct use from UserTokenMiddleware (fallback)
    """
    from agent_utilities.mcp.delegated_auth import (
        get_delegated_token,
        get_user_identity,
        get_user_token,
        is_delegation_enabled,
    )

    CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
    AUTHORITY = "https://login.microsoftonline.com/common"
    SCOPES = [
        "User.Read",
        "Mail.ReadWrite",
        "Calendars.ReadWrite",
        "Files.ReadWrite",
    ]

    # --- Path 1: OIDC Delegation (RFC 8693 Token Exchange) ---
    if is_delegation_enabled():
        try:
            delegated_token = get_delegated_token(
                audience=os.environ.get("AUDIENCE", "https://graph.microsoft.com"),
                scopes=os.environ.get("DELEGATED_SCOPES", " ".join(SCOPES)),
            )
            identity = get_user_identity()
            logger.info(
                "Using OIDC delegated token for Microsoft Graph API",
                extra={"user_email": identity.get("email")},
            )
            # Create AuthManager and inject the delegated token
            auth = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)
            auth.access_token = delegated_token
            return MicrosoftGraphApi(auth)
        except Exception as e:
            logger.warning(f"OIDC delegation failed, falling back to MSAL: {e}")

    # --- Path 2: MSAL Token Cache (silent acquisition) ---
    auth = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)
    token = auth.get_token()

    if token:
        logger.info("Using cached MSAL token for Microsoft Graph API")
        try:
            return MicrosoftGraphApi(auth)
        except (AuthError, UnauthorizedError) as e:
            logger.warning(f"Cached MSAL token failed: {e}")

    # --- Path 3: MCP User Token (direct passthrough) ---
    user_token = get_user_token()
    if user_token:
        logger.info("Using MCP user token passthrough for Microsoft Graph API")
        auth.access_token = user_token
        try:
            return MicrosoftGraphApi(auth)
        except (AuthError, UnauthorizedError) as e:
            raise RuntimeError(
                f"AUTHENTICATION ERROR: The Microsoft credentials are not valid. "
                f"Error details: {str(e)}"
            ) from e

    raise ValueError(
        "Microsoft token is not provided. Please login via device code flow, "
        "configure OIDC delegation, or ensure a valid MCP user token is available."
    )
