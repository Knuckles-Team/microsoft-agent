from unittest.mock import MagicMock, mock_open, patch

import pytest
from agent_utilities.exceptions import AuthError
from keyring.errors import KeyringError

from microsoft_agent.auth import AuthManager, get_client


@pytest.fixture
def mock_msal():
    with (
        patch("msal.PublicClientApplication") as mock_app_cls,
        patch("msal.SerializableTokenCache") as mock_cache_cls,
        patch("atexit.register"),
    ):
        mock_app = MagicMock()
        mock_cache = MagicMock()
        mock_app_cls.return_value = mock_app
        mock_cache_cls.return_value = mock_cache

        yield mock_app, mock_cache


@patch("microsoft_agent.auth.keyring")
@patch("microsoft_agent.auth.FALLBACK_PATH")
@patch("microsoft_agent.auth.SELECTED_ACCOUNT_PATH")
@pytest.mark.concept("ECO-4.1")
def test_auth_manager_init_load_cache(
    mock_selected_path, mock_fallback_path, mock_keyring, mock_msal
):
    mock_app, mock_cache = mock_msal

    # Set up keyring return values
    mock_keyring.get_password.side_effect = [
        "cache_serialized_data",
        '{"account_id": "acc_123"}',
    ]

    # Initialize
    manager = AuthManager("client_id", "authority", ["scope"])

    assert manager.client_id == "client_id"
    assert manager.authority == "authority"
    assert manager.scopes == ["scope"]
    assert manager.selected_account_id == "acc_123"
    mock_cache.deserialize.assert_called_once_with("cache_serialized_data")


@patch("microsoft_agent.auth.keyring")
@patch("microsoft_agent.auth.FALLBACK_PATH")
@patch("microsoft_agent.auth.SELECTED_ACCOUNT_PATH")
@pytest.mark.concept("ECO-4.1")
def test_auth_manager_file_fallback(
    mock_selected_path, mock_fallback_path, mock_keyring, mock_msal
):
    mock_app, mock_cache = mock_msal

    # Keyring fails, files exist
    mock_keyring.get_password.side_effect = KeyringError("keyring error")
    mock_fallback_path.exists.return_value = True
    mock_selected_path.exists.return_value = True

    # Mock open for fallback and selected account
    m_open = mock_open()
    m_open.side_effect = [
        mock_open(read_data="file_serialized_data").return_value,
        mock_open(read_data='{"account_id": "file_acc"}').return_value,
    ]

    with patch("builtins.open", m_open):
        manager = AuthManager("client_id", "authority", ["scope"])

    assert manager.selected_account_id == "file_acc"
    mock_cache.deserialize.assert_called_once_with("file_serialized_data")


@patch("microsoft_agent.auth.keyring")
@patch("microsoft_agent.auth.FALLBACK_PATH")
@patch("microsoft_agent.auth.SELECTED_ACCOUNT_PATH")
@pytest.mark.concept("ECO-4.1")
def test_auth_manager_save_cache(
    mock_selected_path, mock_fallback_path, mock_keyring, mock_msal
):
    mock_app, mock_cache = mock_msal
    mock_keyring.get_password.return_value = None

    manager = AuthManager("client_id", "authority", ["scope"])

    # Trigger save
    mock_cache.has_state_changed = True
    mock_cache.serialize.return_value = "new_cache_data"
    manager.selected_account_id = "new_acc"

    # Keyring works
    manager.save_token_cache()
    mock_keyring.set_password.assert_any_call(
        "microsoft-agent-mcp", "msal_token_cache", "new_cache_data"
    )
    mock_keyring.set_password.assert_any_call(
        "microsoft-agent-mcp", "selected_account", '{"account_id": "new_acc"}'
    )

    # Keyring fails, fallback to files
    mock_keyring.set_password.side_effect = KeyringError("Keyring fail")
    m_open = mock_open()
    with patch("builtins.open", m_open):
        manager.save_token_cache()

    m_open.assert_any_call(mock_fallback_path, "w")
    m_open.assert_any_call(mock_selected_path, "w")


@pytest.mark.concept("ECO-4.1")
def test_get_current_account(mock_msal):
    mock_app, _ = mock_msal
    manager = AuthManager("client_id", "authority", ["scope"])

    # No accounts
    mock_app.get_accounts.return_value = []
    assert manager.get_current_account() is None

    # Accounts, but none match selected
    acc1 = {"home_account_id": "1"}
    acc2 = {"home_account_id": "2"}
    mock_app.get_accounts.return_value = [acc1, acc2]
    manager.selected_account_id = "3"
    assert manager.get_current_account() == acc1

    # Selected matches
    manager.selected_account_id = "2"
    assert manager.get_current_account() == acc2


@pytest.mark.concept("ECO-4.1")
def test_get_token_and_details(mock_msal):
    mock_app, _ = mock_msal
    manager = AuthManager("client_id", "authority", ["scope"])

    acc = {"home_account_id": "1"}
    mock_app.get_accounts.return_value = [acc]

    # acquire_token_silent succeeds
    mock_app.acquire_token_silent.return_value = {"access_token": "token123"}
    assert manager.get_token() == "token123"

    # acquire_token_silent returns none/fails
    mock_app.acquire_token_silent.return_value = None
    assert manager.get_token() is None

    # get_token_details
    mock_app.acquire_token_silent.return_value = {
        "access_token": "token123",
        "expires_in": 3600,
    }
    res = manager.get_token_details(tenant_id="tenant_abc")
    assert res["access_token"] == "token123"
    mock_app.acquire_token_silent.assert_called_with(
        ["scope"],
        account=acc,
        claims_challenge=None,
        authority="https://login.microsoftonline.com/tenant_abc",
    )


@pytest.mark.concept("ECO-4.1")
def test_acquire_token_by_device_code(mock_msal):
    mock_app, _ = mock_msal
    manager = AuthManager("client_id", "authority", ["scope"])

    # 1. Flow init fails
    mock_app.initiate_device_flow.return_value = {}
    with pytest.raises(Exception, match="Failed to create device flow"):
        manager.acquire_token_by_device_code(lambda _: None)

    # 2. Flow success
    flow = {"user_code": "CODE", "message": "Go to link"}
    mock_app.initiate_device_flow.return_value = flow
    mock_app.acquire_token_by_device_flow.return_value = {
        "access_token": "token_device",
        "home_account_id": "acc_device",
    }

    callback = MagicMock()
    with (
        patch.object(manager, "save_token_cache") as mock_save_cache,
        patch.object(manager, "save_selected_account") as mock_save_selected,
    ):
        msg = manager.acquire_token_by_device_code(callback)
        assert msg == "Authentication successful"
        assert manager.access_token == "token_device"
        assert manager.selected_account_id == "acc_device"
        callback.assert_called_once_with("Go to link")
        mock_save_cache.assert_called_once()
        mock_save_selected.assert_called_once()

    # 3. Flow failure response
    mock_app.acquire_token_by_device_flow.return_value = {
        "error_description": "User expired"
    }
    with pytest.raises(Exception, match="Authentication failed: User expired"):
        manager.acquire_token_by_device_code(callback)


@patch("microsoft_agent.auth.keyring")
@patch("microsoft_agent.auth.FALLBACK_PATH")
@patch("microsoft_agent.auth.SELECTED_ACCOUNT_PATH")
@pytest.mark.concept("ECO-4.1")
def test_logout(mock_selected_path, mock_fallback_path, mock_keyring, mock_msal):
    mock_app, _ = mock_msal
    mock_keyring.get_password.return_value = None
    manager = AuthManager("client_id", "authority", ["scope"])

    acc = {"home_account_id": "1"}
    mock_app.get_accounts.return_value = [acc]
    mock_fallback_path.exists.return_value = True
    mock_selected_path.exists.return_value = True

    manager.logout()
    mock_app.remove_account.assert_called_once_with(acc)
    assert manager.selected_account_id is None
    assert manager.access_token is None
    mock_keyring.delete_password.assert_any_call(
        "microsoft-agent-mcp", "msal_token_cache"
    )
    mock_keyring.delete_password.assert_any_call(
        "microsoft-agent-mcp", "selected_account"
    )
    mock_fallback_path.unlink.assert_called_once()
    mock_selected_path.unlink.assert_called_once()


@pytest.mark.concept("ECO-4.1")
def test_list_select_remove_account(mock_msal):
    mock_app, _ = mock_msal
    manager = AuthManager("client_id", "authority", ["scope"])

    acc1 = {"home_account_id": "1"}
    acc2 = {"home_account_id": "2"}
    mock_app.get_accounts.return_value = [acc1, acc2]

    # list_accounts
    assert manager.list_accounts() == [acc1, acc2]

    # select_account
    with patch.object(manager, "save_selected_account") as mock_save:
        assert manager.select_account("2") is True
        assert manager.selected_account_id == "2"
        mock_save.assert_called_once()

        assert manager.select_account("nonexistent") is False

    # remove_account
    with (
        patch.object(manager, "save_selected_account") as mock_save_sel,
        patch.object(manager, "save_token_cache") as mock_save_cache,
    ):
        manager.selected_account_id = "2"
        assert manager.remove_account("2") is True
        mock_app.remove_account.assert_called_once_with(acc2)
        assert manager.selected_account_id is None
        mock_save_sel.assert_called_once()
        mock_save_cache.assert_called_once()


@pytest.mark.asyncio
@patch("agent_utilities.mcp.delegated_auth.is_delegation_enabled")
@patch("agent_utilities.mcp.delegated_auth.get_delegated_token")
@patch("agent_utilities.mcp.delegated_auth.get_user_identity")
@patch("agent_utilities.mcp.delegated_auth.get_user_token")
@patch("microsoft_agent.auth.AuthManager")
@patch("microsoft_agent.api_client.MicrosoftGraphApi")
@pytest.mark.concept("ECO-4.1")
async def test_get_client_flows(
    mock_api_cls,
    mock_auth_manager_cls,
    mock_get_user_token,
    mock_get_user_identity,
    mock_get_delegated_token,
    mock_is_delegation_enabled,
):
    # 1. Path 1: Delegation is enabled and succeeds
    mock_is_delegation_enabled.return_value = True
    mock_get_delegated_token.return_value = "del_token"
    mock_get_user_identity.return_value = {"email": "user@org.com"}

    mock_auth_instance = MagicMock()
    mock_auth_manager_cls.return_value = mock_auth_instance

    await get_client()
    mock_api_cls.assert_called_once_with(mock_auth_instance)
    assert mock_auth_instance.access_token == "del_token"

    # Reset
    mock_api_cls.reset_mock()
    mock_auth_manager_cls.reset_mock()

    # 2. Path 1: Delegation is enabled but raises exception, falls back to silent MSAL which succeeds
    mock_is_delegation_enabled.return_value = True
    mock_get_delegated_token.side_effect = Exception("delegation failed")
    mock_auth_instance.get_token.return_value = "silent_token"

    await get_client()
    mock_api_cls.assert_called_once_with(mock_auth_instance)

    # Reset
    mock_api_cls.reset_mock()
    mock_auth_manager_cls.reset_mock()

    # 3. Path 2: Silent MSAL fails, falls back to Path 3: MCP User Token which succeeds
    mock_is_delegation_enabled.return_value = False
    mock_auth_instance.get_token.return_value = None
    mock_get_user_token.return_value = "mcp_user_token"

    await get_client()
    assert mock_auth_instance.access_token == "mcp_user_token"
    mock_api_cls.assert_called_once_with(mock_auth_instance)

    # Reset
    mock_api_cls.reset_mock()
    mock_auth_manager_cls.reset_mock()

    # 4. Path 3: MCP User Token fails with AuthError -> raises RuntimeError
    mock_is_delegation_enabled.return_value = False
    mock_auth_instance.get_token.return_value = None
    mock_get_user_token.return_value = "invalid_user_token"
    mock_api_cls.side_effect = AuthError("Invalid credential")

    with pytest.raises(
        RuntimeError,
        match="AUTHENTICATION ERROR: The Microsoft credentials are not valid",
    ):
        await get_client()

    # Reset side effect
    mock_api_cls.side_effect = None

    # 5. Path 3: No token available -> raises ValueError
    mock_is_delegation_enabled.return_value = False
    mock_auth_instance.get_token.return_value = None
    mock_get_user_token.return_value = None

    with pytest.raises(ValueError, match="Microsoft token is not provided"):
        await get_client()
