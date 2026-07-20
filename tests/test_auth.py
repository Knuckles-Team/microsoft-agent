"""Tests for auth.py module."""

import base64
import json
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from keyring.errors import KeyringError

from microsoft_agent.auth import (
    SELECTED_ACCOUNT_KEY,
    SERVICE_NAME,
    TOKEN_CACHE_ACCOUNT,
    AuthManager,
)
from microsoft_agent.settings import AuthenticationMode


class TestAuthManager:
    """Test AuthManager class."""

    def test_init(self, real_auth_manager):
        """Test AuthManager initialization."""
        assert real_auth_manager.client_id == "test_client_id"
        assert real_auth_manager.authority == "https://login.microsoftonline.com/common"
        assert real_auth_manager.scopes == ["User.Read", "Mail.ReadWrite"]
        assert real_auth_manager.access_token is None

    def test_managed_identity_credential_acquires_default_scope(self):
        credential = MagicMock()
        credential.get_token.return_value = MagicMock(
            token="managed-token", expires_on=2_000_000_000
        )
        with patch("microsoft_agent.auth.keyring.get_password") as keyring_get:
            manager = AuthManager(
                "",
                "https://login.microsoftonline.com/organizations",
                ["https://graph.microsoft.com/.default"],
                mode=AuthenticationMode.MANAGED_IDENTITY,
                azure_credential=credential,
            )

        details = manager.get_token_details()

        assert details == {
            "access_token": "managed-token",
            "expires_on": 2_000_000_000,
        }
        credential.get_token.assert_called_once_with(
            "https://graph.microsoft.com/.default"
        )
        keyring_get.assert_not_called()

    def test_external_graph_token_is_not_reused_for_another_audience(self):
        payload = base64.urlsafe_b64encode(
            json.dumps(
                {
                    "tid": "tenant-id",
                    "aud": "https://graph.microsoft.com",
                    "exp": 2_000_000_000,
                }
            ).encode()
        ).rstrip(b"=")
        token = f"header.{payload.decode()}.signature"
        manager = AuthManager(
            "client-id",
            "https://login.microsoftonline.com/tenant-id",
            ["https://graph.microsoft.com/.default"],
            mode=AuthenticationMode.EXTERNAL_TOKEN,
            external_token_provider=lambda: token,
            expected_tenant_id="tenant-id",
            allowed_external_audiences=("https://graph.microsoft.com",),
        )

        assert (
            manager.acquire_token_for_scopes(["api://windows-companion/.default"])
            is None
        )
        assert manager.get_token(["https://graph.microsoft.com/.default"]) == token

    def test_load_token_cache_from_keyring(self, real_auth_manager):
        """Test loading token cache from keyring."""
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.return_value = '{"test": "data"}'
            real_auth_manager.load_token_cache()
            # Should be called twice: once for token cache, once for selected account
            assert mock_get.call_count == 2
            mock_get.assert_any_call(SERVICE_NAME, TOKEN_CACHE_ACCOUNT)

    def test_load_token_cache_keyring_error(self, real_auth_manager):
        """A keyring failure does not enable plaintext token persistence."""
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.side_effect = KeyringError("Keyring error")
            real_auth_manager.load_token_cache()

    def test_load_token_cache_no_cache(self, real_auth_manager):
        """Test loading when no cache exists."""
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.return_value = None
            real_auth_manager.load_token_cache()

    def test_save_token_cache_to_keyring(self, real_auth_manager):
        """Test saving token cache to keyring."""
        real_auth_manager.token_cache.has_state_changed = True
        real_auth_manager.selected_account_id = (
            "test_account"  # Set an account to ensure both saves happen
        )
        with patch.object(
            real_auth_manager.token_cache, "serialize", return_value='{"test": "data"}'
        ):
            with patch("microsoft_agent.auth.keyring.set_password") as mock_set:
                real_auth_manager.save_token_cache()
                # Should be called twice: once for token cache, once for selected account
                assert mock_set.call_count == 2
                mock_set.assert_any_call(
                    SERVICE_NAME, TOKEN_CACHE_ACCOUNT, '{"test": "data"}'
                )

    def test_save_token_cache_keyring_error(self, real_auth_manager):
        """A keyring failure leaves the token cache in memory only."""
        real_auth_manager.token_cache.has_state_changed = True
        with patch.object(
            real_auth_manager.token_cache, "serialize", return_value='{"test": "data"}'
        ):
            with patch("microsoft_agent.auth.keyring.set_password") as mock_set:
                mock_set.side_effect = KeyringError("Keyring error")
                real_auth_manager.save_token_cache()

    def test_save_token_cache_no_change(self, real_auth_manager):
        """Test saving when cache hasn't changed."""
        real_auth_manager.token_cache.has_state_changed = False
        real_auth_manager.selected_account_id = None  # Ensure no account to save

        with patch("microsoft_agent.auth.keyring.set_password") as mock_set:
            real_auth_manager.save_token_cache()
            # Should not be called since cache hasn't changed and no selected account
            mock_set.assert_not_called()

    def test_load_selected_account_from_keyring(self, real_auth_manager):
        """Test loading selected account from keyring."""
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.return_value = '{"account_id": "account123"}'
            real_auth_manager.load_selected_account()
            assert real_auth_manager.selected_account_id == "account123"

    def test_load_selected_account_keyring_error(self, real_auth_manager):
        """A keyring failure does not read account identifiers from disk."""
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.side_effect = KeyringError("Keyring error")
            real_auth_manager.load_selected_account()
            assert real_auth_manager.selected_account_id is None

    def test_load_selected_account_invalid_json(self, real_auth_manager):
        """Test loading selected account with invalid JSON."""
        real_auth_manager.selected_account_id = (
            "existing_id"  # Start with an existing ID
        )
        with patch("microsoft_agent.auth.keyring.get_password") as mock_get:
            mock_get.return_value = "invalid json"
            real_auth_manager.load_selected_account()
            # Should remain unchanged due to invalid JSON
            assert real_auth_manager.selected_account_id == "existing_id"

    def test_save_selected_account(self, real_auth_manager):
        """Test saving selected account."""
        real_auth_manager.selected_account_id = "account789"

        with patch("microsoft_agent.auth.keyring.set_password") as mock_set:
            real_auth_manager.save_selected_account()
            expected_data = json.dumps({"account_id": "account789"})
            mock_set.assert_called_once_with(
                SERVICE_NAME, SELECTED_ACCOUNT_KEY, expected_data
            )

    def test_save_selected_account_none(self, real_auth_manager):
        """Test saving when selected account is None."""
        real_auth_manager.selected_account_id = None

        with patch("microsoft_agent.auth.keyring.set_password") as mock_set:
            real_auth_manager.save_selected_account()
            mock_set.assert_not_called()

    def test_get_current_account_no_accounts(self, real_auth_manager):
        """Test getting current account when no accounts exist."""
        real_auth_manager.msal_app.get_accounts.return_value = []
        result = real_auth_manager.get_current_account()
        assert result is None

    def test_get_current_account_with_selected_account(
        self, real_auth_manager, mock_account
    ):
        """Test getting current account with selected account."""
        real_auth_manager.selected_account_id = "account123"
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.get_current_account()
        assert result == mock_account

    def test_get_current_account_selected_not_found(
        self, real_auth_manager, mock_account
    ):
        """Test getting current account when selected account not found."""
        real_auth_manager.selected_account_id = "wrong_id"
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.get_current_account()
        assert result is None

    def test_get_current_account_no_selected(self, real_auth_manager, mock_account):
        """Test getting current account when no account selected."""
        real_auth_manager.selected_account_id = None
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.get_current_account()
        assert result == mock_account

    def test_get_token_no_account(self, real_auth_manager):
        """Test getting token when no account exists."""
        real_auth_manager.msal_app.get_accounts.return_value = []
        result = real_auth_manager.get_token()
        assert result is None

    def test_get_token_success(
        self, real_auth_manager, mock_account, mock_token_response
    ):
        """Test getting token successfully."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.msal_app.acquire_token_silent.return_value = (
            mock_token_response
        )

        result = real_auth_manager.get_token()
        assert result == "test_access_token"

    def test_get_token_no_access_token(self, real_auth_manager, mock_account):
        """Test getting token when response has no access token."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.msal_app.acquire_token_silent.return_value = {}

        result = real_auth_manager.get_token()
        assert result is None

    def test_get_token_details_success(
        self, real_auth_manager, mock_account, mock_token_response
    ):
        """Test getting token details successfully."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.msal_app.acquire_token_silent.return_value = (
            mock_token_response
        )

        result = real_auth_manager.get_token_details()
        assert result == mock_token_response

    def test_get_token_details_no_account(self, real_auth_manager):
        """Test getting token details when no account exists."""
        real_auth_manager.msal_app.get_accounts.return_value = []
        result = real_auth_manager.get_token_details()
        assert result is None

    def test_acquire_token_by_device_code_success(
        self, real_auth_manager, mock_device_flow, mock_token_response
    ):
        """Test device code flow success."""
        real_auth_manager.allow_device_code = True
        real_auth_manager.msal_app.initiate_device_flow.return_value = mock_device_flow
        real_auth_manager.msal_app.acquire_token_by_device_flow.return_value = (
            mock_token_response
        )

        callback = MagicMock()
        result = real_auth_manager.acquire_token_by_device_code(callback)

        assert result == "Authentication successful"
        assert real_auth_manager.access_token == "test_access_token"
        assert real_auth_manager.selected_account_id == "account123"
        callback.assert_called_once()

    def test_acquire_token_by_device_code_failure(
        self, real_auth_manager, mock_device_flow
    ):
        """Test device code flow failure."""
        real_auth_manager.allow_device_code = True
        mock_device_flow["user_code"] = "TEST123"  # Ensure user_code exists
        real_auth_manager.msal_app.initiate_device_flow.return_value = mock_device_flow
        real_auth_manager.msal_app.acquire_token_by_device_flow.return_value = {
            "error": "access_denied",
            "error_description": "User denied access",
        }

        callback = MagicMock()
        with pytest.raises(Exception, match="Authentication failed"):
            real_auth_manager.acquire_token_by_device_code(callback)

    def test_acquire_token_by_device_code_no_user_code(self, real_auth_manager):
        """Test device code flow when user_code is missing."""
        real_auth_manager.allow_device_code = True
        real_auth_manager.msal_app.initiate_device_flow.return_value = {}

        callback = MagicMock()
        with pytest.raises(Exception, match="Failed to create device flow"):
            real_auth_manager.acquire_token_by_device_code(callback)

    def test_device_code_is_disabled_by_default(self, real_auth_manager):
        with pytest.raises(ValueError, match="disabled"):
            real_auth_manager.acquire_token_by_device_code(MagicMock())

        real_auth_manager.msal_app.initiate_device_flow.assert_not_called()

    def test_interactive_login_fails_when_required_keyring_is_unavailable(
        self, real_auth_manager
    ):
        real_auth_manager.secure_cache_available = False
        real_auth_manager.require_secure_cache = True

        with pytest.raises(Exception, match="Secure OS token storage"):
            real_auth_manager.acquire_token_interactive()

        real_auth_manager.msal_app.acquire_token_interactive.assert_not_called()

    def test_logout(self, real_auth_manager, mock_account):
        """Test logout."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.selected_account_id = "account123"
        real_auth_manager.access_token = "test_token"

        with patch("microsoft_agent.auth.keyring.delete_password"):
            real_auth_manager.logout()

            assert real_auth_manager.selected_account_id is None
            assert real_auth_manager.access_token is None
            real_auth_manager.msal_app.remove_account.assert_called_once_with(
                mock_account
            )

    def test_logout_keyring_error(self, real_auth_manager, mock_account):
        """Test logout when keyring delete fails."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        with patch("microsoft_agent.auth.keyring.delete_password") as mock_delete:
            mock_delete.side_effect = Exception("Delete failed")

            # Should not raise exception
            real_auth_manager.logout()

    def test_list_accounts(self, real_auth_manager, mock_account):
        """Test listing accounts."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        result = real_auth_manager.list_accounts()
        assert result == [mock_account]

    def test_select_account_success(self, real_auth_manager, mock_account):
        """Test selecting an account successfully."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.select_account("account123")
        assert result is True
        assert real_auth_manager.selected_account_id == "account123"

    def test_select_account_not_found(self, real_auth_manager, mock_account):
        """Test selecting an account that doesn't exist."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.select_account("wrong_id")
        assert result is False

    def test_remove_account_success(self, real_auth_manager, mock_account):
        """Test removing an account successfully."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.selected_account_id = "account123"

        result = real_auth_manager.remove_account("account123")
        assert result is True
        assert real_auth_manager.selected_account_id is None

    def test_remove_account_not_found(self, real_auth_manager, mock_account):
        """Test removing an account that doesn't exist."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]

        result = real_auth_manager.remove_account("wrong_id")
        assert result is False

    def test_remove_account_selected_different(self, real_auth_manager, mock_account):
        """Test removing an account when selected account is different."""
        real_auth_manager.msal_app.get_accounts.return_value = [mock_account]
        real_auth_manager.selected_account_id = "different_id"

        result = real_auth_manager.remove_account("account123")
        assert result is True
        assert (
            real_auth_manager.selected_account_id == "different_id"
        )  # Should remain unchanged


@pytest.mark.asyncio
async def test_get_client_with_token():
    """Test get_client with valid token."""
    mock_auth = MagicMock()
    mock_auth.get_token.return_value = "test_token"
    mock_auth.scopes = ["User.Read"]
    with patch("microsoft_agent.auth.get_auth_manager", return_value=mock_auth):
        with patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api:
            from microsoft_agent.auth import get_client

            result = await get_client()
            assert result == mock_api.return_value


@pytest.mark.asyncio
async def test_graph_client_dependency_closes_request_transport():
    """The FastMCP dependency owns and closes its request-scoped client."""

    client = MagicMock()
    client.close = AsyncMock()
    with patch("microsoft_agent.auth.get_client", AsyncMock(return_value=client)):
        from microsoft_agent.auth import get_client_dependency

        dependency = get_client_dependency()
        assert await anext(dependency) is client
        await dependency.aclose()

    client.close.assert_awaited_once_with()


@pytest.mark.asyncio
async def test_get_client_without_token():
    """Test get_client without token raises error."""
    mock_auth = MagicMock()
    mock_auth.get_token.return_value = None
    with patch("microsoft_agent.auth.get_auth_manager", return_value=mock_auth):
        from microsoft_agent.auth import get_client

        with pytest.raises(ValueError, match="Microsoft token is not available"):
            await get_client()


@pytest.mark.asyncio
async def test_get_client_auth_error():
    """Test get_client with authentication error."""
    from agent_utilities.core.exceptions import AuthError

    mock_auth = MagicMock()
    mock_auth.get_token.return_value = "test_token"
    with patch("microsoft_agent.auth.get_auth_manager", return_value=mock_auth):
        with patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api:
            mock_api.side_effect = AuthError("Auth failed")

            from microsoft_agent.auth import get_client

            with pytest.raises(RuntimeError, match="AUTHENTICATION ERROR"):
                await get_client()
