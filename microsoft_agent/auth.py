import json
import logging
import atexit
from typing import List, Optional, Dict, Any
import msal
import keyring
from keyring.errors import KeyringError
from pathlib import Path

logger = logging.getLogger(__name__)

SERVICE_NAME = "microsoft-agent-mcp"
TOKEN_CACHE_ACCOUNT = "msal_token_cache"
SELECTED_ACCOUNT_KEY = "selected_account"
FALLBACK_DIR = Path.home() / ".microsoft-agent"
FALLBACK_PATH = FALLBACK_DIR / ".token_cache.json"
SELECTED_ACCOUNT_PATH = FALLBACK_DIR / ".selected_account.json"

FALLBACK_DIR.mkdir(parents=True, exist_ok=True)


class AuthManager:
    def __init__(self, client_id: str, authority: str, scopes: List[str]):
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes
        self.token_cache = msal.SerializableTokenCache()
        self.msal_app = msal.PublicClientApplication(
            self.client_id, authority=self.authority, token_cache=self.token_cache
        )
        self.access_token: Optional[str] = None
        self.selected_account_id: Optional[str] = None

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
                with open(FALLBACK_PATH, "r") as f:
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
                with open(SELECTED_ACCOUNT_PATH, "r") as f:
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

    def get_current_account(self) -> Optional[Dict[str, Any]]:
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

    def get_token(self) -> Optional[str]:
        account = self.get_current_account()
        if not account:
            return None

        result = self.msal_app.acquire_token_silent(self.scopes, account=account)
        if result and "access_token" in result:
            return result["access_token"]

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
        except Exception:
            pass

        if FALLBACK_PATH.exists():
            FALLBACK_PATH.unlink()
        if SELECTED_ACCOUNT_PATH.exists():
            SELECTED_ACCOUNT_PATH.unlink()

    def list_accounts(self) -> List[Dict[str, Any]]:
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
