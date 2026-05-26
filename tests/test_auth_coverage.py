import os
import sys
import json
import time
import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
from azure.core.credentials import AccessToken

from agent_utilities.exceptions import AuthError, UnauthorizedError
import microsoft_agent.auth as auth_mod
from microsoft_agent.auth import AuthManager, get_client
from microsoft_agent.credential_adapter import AuthManagerCredential


@pytest.fixture
def mock_fallback_paths(tmp_path, monkeypatch):
    """Redirect token cache files to temporary directory to prevent system file pollution."""
    fb_dir = tmp_path / "ms_agent"
    fb_dir.mkdir(parents=True, exist_ok=True)
    fb_path = fb_dir / ".token_cache.json"
    sel_path = fb_dir / ".selected_account.json"

    monkeypatch.setattr(auth_mod, "FALLBACK_DIR", fb_dir)
    monkeypatch.setattr(auth_mod, "FALLBACK_PATH", fb_path)
    monkeypatch.setattr(auth_mod, "SELECTED_ACCOUNT_PATH", sel_path)

    return fb_dir, fb_path, sel_path


@pytest.fixture
def mock_keyring(monkeypatch):
    """Mock keyring to isolate tests from host keyring."""
    store = {}

    def get_password(service, username):
        if service == "fail":
            from keyring.errors import KeyringError

            raise KeyringError("Keyring error")
        if service == "import_fail":
            raise ImportError("Keyring not installed")
        return store.get((service, username))

    def set_password(service, username, password):
        if service == "fail":
            from keyring.errors import KeyringError

            raise KeyringError("Keyring error")
        if service == "import_fail":
            raise ImportError("Keyring not installed")
        store[(service, username)] = password

    def delete_password(service, username):
        if service == "fail":
            raise Exception("Keyring delete failure")
        store.pop((service, username), None)

    monkeypatch.setattr(auth_mod.keyring, "get_password", get_password)
    monkeypatch.setattr(auth_mod.keyring, "set_password", set_password)
    monkeypatch.setattr(auth_mod.keyring, "delete_password", delete_password)

    return store


@pytest.fixture
def mock_msal(monkeypatch):
    """Mock MSAL library methods."""
    mock_app_instance = MagicMock()
    mock_app_class = MagicMock(return_value=mock_app_instance)

    mock_cache_instance = MagicMock()
    mock_cache_class = MagicMock(return_value=mock_cache_instance)

    monkeypatch.setattr(auth_mod.msal, "PublicClientApplication", mock_app_class)
    monkeypatch.setattr(auth_mod.msal, "SerializableTokenCache", mock_cache_class)
    monkeypatch.setattr(auth_mod.atexit, "register", lambda fn: None)
    return mock_app_instance


@pytest.mark.concept("ECO-4.1")
def test_auth_manager_init_and_cache_loading(
    mock_fallback_paths, mock_keyring, mock_msal
):
    # Test keyring fallback to file
    mock_keyring[("fail", "msal_token_cache")] = "keyring_data"  # Trigger service fail
    # Set fallback service name temporarily to trigger keyring fail path
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        # Create AuthManager
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.client_id == "client_id"
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_load_token_cache_from_file(mock_fallback_paths, mock_keyring, mock_msal):
    _, fb_path, _ = mock_fallback_paths
    fb_path.write_text("file_cache_data")

    # Force keyring fail to hit file fallback
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.token_cache.deserialize.called
        # Check argument passed to deserialize
        auth.token_cache.deserialize.assert_called_with("file_cache_data")
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_load_token_cache_file_read_error(mock_fallback_paths, mock_keyring, mock_msal):
    _, fb_path, _ = mock_fallback_paths
    fb_path.write_text("dummy")

    # Force keyring fail to hit file fallback
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        with patch("builtins.open", side_effect=OSError("Read error")):
            auth = AuthManager("client_id", "authority", ["User.Read"])
            # Should catch OSError in load_token_cache and not raise
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_save_token_cache(mock_fallback_paths, mock_keyring, mock_msal):
    _, fb_path, _ = mock_fallback_paths
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Mock token cache state change
    auth.token_cache.has_state_changed = True
    auth.token_cache.serialize.return_value = "new_cache_data"

    # Save to keyring
    auth.save_token_cache()
    assert (
        mock_keyring.get((auth_mod.SERVICE_NAME, auth_mod.TOKEN_CACHE_ACCOUNT))
        == "new_cache_data"
    )

    # Mock keyring failure on save to test file writing fallback
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        auth.save_token_cache()
        assert fb_path.read_text() == "new_cache_data"
    finally:
        auth_mod.SERVICE_NAME = original_service

    # Test file writing failure
    auth_mod.SERVICE_NAME = "fail"
    try:
        with patch("builtins.open", side_effect=OSError("Write error")):
            auth.save_token_cache()  # Should catch OSError in save_token_cache and not raise
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_load_selected_account(mock_fallback_paths, mock_keyring, mock_msal):
    _, _, sel_path = mock_fallback_paths
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        # Invalid JSON fallback
        sel_path.write_text("invalid json")
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.selected_account_id is None

        # Valid JSON fallback
        sel_path.write_text('{"account_id": "test_account_id"}')
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.selected_account_id == "test_account_id"

        # OSError on load selected account
        with patch("builtins.open", side_effect=OSError("Read error")):
            auth = AuthManager("client_id", "authority", ["User.Read"])
            # Should catch OSError and not raise
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_save_selected_account(mock_fallback_paths, mock_keyring, mock_msal):
    _, _, sel_path = mock_fallback_paths
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No selected account
    auth.selected_account_id = None
    auth.save_selected_account()
    assert not sel_path.exists()

    # Case 2: Save to keyring
    auth.selected_account_id = "keyring_acc"
    auth.save_selected_account()
    assert "keyring_acc" in mock_keyring.get(
        (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY)
    )

    # Case 3: Keyring fail fallback to file
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth.selected_account_id = "file_acc"
        auth.save_selected_account()
        assert "file_acc" in sel_path.read_text()
    finally:
        auth_mod.SERVICE_NAME = original_service

    # Case 4: File write failure
    auth_mod.SERVICE_NAME = "fail"
    try:
        with patch("builtins.open", side_effect=OSError("Write error")):
            auth.save_selected_account()  # Should not raise exception
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_get_current_account(mock_fallback_paths, mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No accounts
    mock_msal.get_accounts.return_value = []
    assert auth.get_current_account() is None

    # Case 2: Has accounts, but no selected account (returns first)
    acc1 = {"home_account_id": "acc1"}
    acc2 = {"home_account_id": "acc2"}
    mock_msal.get_accounts.return_value = [acc1, acc2]
    auth.selected_account_id = None
    assert auth.get_current_account() == acc1

    # Case 3: Has selected account in list
    auth.selected_account_id = "acc2"
    assert auth.get_current_account() == acc2

    # Case 4: Selected account not in list (warns and returns first)
    auth.selected_account_id = "missing"
    assert auth.get_current_account() == acc1


@pytest.mark.concept("ECO-4.1")
def test_get_token(mock_fallback_paths, mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No account
    mock_msal.get_accounts.return_value = []
    assert auth.get_token() is None

    # Case 2: Account exists, acquire silent returns token
    acc = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc]
    mock_msal.acquire_token_silent.return_value = {"access_token": "token123"}
    assert auth.get_token() == "token123"

    # Case 3: Acquire silent returns None
    mock_msal.acquire_token_silent.return_value = None
    assert auth.get_token() is None


@pytest.mark.concept("ECO-4.1")
def test_get_token_details(mock_fallback_paths, mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No account
    mock_msal.get_accounts.return_value = []
    assert auth.get_token_details() is None

    # Case 2: Success with tenant_id authority building
    acc = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc]
    mock_msal.acquire_token_silent.return_value = {
        "access_token": "token123",
        "expires_in": 3600,
    }
    res = auth.get_token_details(claims="some_claim", tenant_id="tenant123")
    assert res["access_token"] == "token123"
    mock_msal.acquire_token_silent.assert_called_with(
        auth.scopes,
        account=acc,
        claims_challenge="some_claim",
        authority="https://login.microsoftonline.com/tenant123",
    )

    # Case 3: Silent return is empty
    mock_msal.acquire_token_silent.return_value = {}
    assert auth.get_token_details() is None


@pytest.mark.concept("ECO-4.1")
def test_acquire_token_by_device_code(mock_fallback_paths, mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: Failed flow init
    mock_msal.initiate_device_flow.return_value = {}
    with pytest.raises(Exception, match="Failed to create device flow"):
        auth.acquire_token_by_device_code(lambda _: None)

    # Case 2: Success flow init, callback called, success return
    mock_msal.initiate_device_flow.return_value = {
        "user_code": "XYZ",
        "message": "Go to link...",
    }
    mock_msal.acquire_token_by_device_flow.return_value = {
        "access_token": "tok456",
        "home_account_id": "acc456",
    }

    callback = MagicMock()
    res = auth.acquire_token_by_device_code(callback)
    callback.assert_called_once_with("Go to link...")
    assert res == "Authentication successful"
    assert auth.access_token == "tok456"
    assert auth.selected_account_id == "acc456"

    # Case 3: Flow failed in execution
    mock_msal.acquire_token_by_device_flow.return_value = {
        "error_description": "User expired"
    }
    with pytest.raises(Exception, match="Authentication failed: User expired"):
        auth.acquire_token_by_device_code(callback)


@pytest.mark.concept("ECO-4.1")
def test_logout_and_account_management(mock_fallback_paths, mock_keyring, mock_msal):
    _, fb_path, sel_path = mock_fallback_paths
    auth = AuthManager("client_id", "authority", ["User.Read"])
    acc1 = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc1]

    # Test list_accounts
    assert auth.list_accounts() == [acc1]

    # Test select_account
    assert not auth.select_account("missing")
    assert auth.select_account("acc1")
    assert auth.selected_account_id == "acc1"

    # Test remove_account
    assert not auth.remove_account("missing")
    assert auth.remove_account("acc1")
    assert auth.selected_account_id is None

    # Test logout with existing files
    auth.token_cache.has_state_changed = True
    auth.token_cache.serialize.return_value = "cache_data"
    auth.selected_account_id = "some_id"
    auth.access_token = "some_token"

    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth.save_token_cache()
        auth.save_selected_account()
    finally:
        auth_mod.SERVICE_NAME = original_service

    assert fb_path.exists()
    assert sel_path.exists()

    auth.logout()
    assert auth.selected_account_id is None
    assert auth.access_token is None
    assert not fb_path.exists()
    assert not sel_path.exists()

    # Test keyring throw and unlinking missing files in logout
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth2 = AuthManager("client_id", "authority", ["User.Read"])
        auth2.logout()  # Should not raise exception
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_oidc_delegation(
    monkeypatch, mock_fallback_paths, mock_keyring, mock_msal
):
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: True
    )
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_delegated_token",
        lambda *args, **kwargs: "oidc_tok",
    )
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_user_identity",
        lambda: {"email": "user@example.com"},
    )

    with patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api:
        client = await get_client()
        assert client is not None
        assert mock_api.called


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_cached_msal_token(
    monkeypatch, mock_fallback_paths, mock_keyring, mock_msal
):
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: False
    )
    acc = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc]
    mock_msal.acquire_token_silent.return_value = {"access_token": "cached_tok"}

    with patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api:
        client = await get_client()
        assert client is not None
        assert mock_api.called


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_user_token_fallback(
    monkeypatch, mock_fallback_paths, mock_keyring, mock_msal
):
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: False
    )
    mock_msal.get_accounts.return_value = []  # No cache
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_user_token",
        lambda: "user_passthrough_tok",
    )

    with patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api:
        client = await get_client()
        assert client is not None
        assert mock_api.called


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_auth_errors_and_fallbacks(
    monkeypatch, mock_fallback_paths, mock_keyring, mock_msal
):
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: False
    )
    mock_msal.get_accounts.return_value = []  # No cache
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_user_token", lambda: None
    )

    # ValueError because no credentials
    with pytest.raises(ValueError, match="Microsoft token is not provided"):
        await get_client()

    # Delegation raises and falls back to ValueError since no other auth
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: True
    )
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_delegated_token",
        MagicMock(side_effect=Exception("Delegation failed")),
    )
    with pytest.raises(ValueError, match="Microsoft token is not provided"):
        await get_client()


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_api_instantiation_errors(
    monkeypatch, mock_fallback_paths, mock_keyring, mock_msal
):
    # MSAL silent path throws AuthError / UnauthorizedError
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.is_delegation_enabled", lambda: False
    )
    acc = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc]
    mock_msal.acquire_token_silent.return_value = {"access_token": "cached_tok"}

    # Mock MicrosoftGraphApi to throw AuthError on instantiation
    with patch(
        "microsoft_agent.api_client.MicrosoftGraphApi",
        side_effect=AuthError("Token invalid"),
    ):
        # Should catch warning and move to Path 3 (raising ValueError in the end because user_token is None)
        with pytest.raises(ValueError, match="Microsoft token is not provided"):
            await get_client()

    # MCP User Token path throws AuthError / UnauthorizedError
    mock_msal.get_accounts.return_value = []  # Clear MSAL cache path
    monkeypatch.setattr(
        "agent_utilities.mcp.delegated_auth.get_user_token",
        lambda: "user_passthrough_tok",
    )

    with patch(
        "microsoft_agent.api_client.MicrosoftGraphApi",
        side_effect=UnauthorizedError("Token expired"),
    ):
        # Should wrap into RuntimeError
        with pytest.raises(
            RuntimeError,
            match="AUTHENTICATION ERROR: The Microsoft credentials are not valid",
        ):
            await get_client()


@pytest.mark.concept("ECO-4.1")
def test_auth_manager_credential_adapter():
    auth_manager = MagicMock()
    adapter = AuthManagerCredential(auth_manager)

    # Case 1: get_token returns successfully when token_details has access_token and expires_on
    auth_manager.get_token_details.return_value = {
        "access_token": "adapter_tok",
        "expires_on": 1234567890,
    }
    tok = adapter.get_token("User.Read")
    assert isinstance(tok, AccessToken)
    assert tok.token == "adapter_tok"
    assert tok.expires_on == 1234567890

    # Case 2: get_token fallback using expires_in when expires_on missing
    auth_manager.get_token_details.return_value = {
        "access_token": "adapter_tok",
        "expires_in": 120,
    }
    tok = adapter.get_token("User.Read")
    assert tok.token == "adapter_tok"
    assert tok.expires_on > time.time()

    # Case 3: get_token fails because token_details is None
    auth_manager.get_token_details.return_value = None
    with pytest.raises(Exception, match="Failed to acquire token"):
        adapter.get_token("User.Read")
