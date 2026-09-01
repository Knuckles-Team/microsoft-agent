import time
from unittest.mock import MagicMock, patch

import pytest
from agent_utilities.core.exceptions import AuthError, UnauthorizedError
from azure.core.credentials import AccessToken

import microsoft_agent.auth as auth_mod
from microsoft_agent.auth import AuthManager, get_client
from microsoft_agent.credential_adapter import AuthManagerCredential


@pytest.fixture
def mock_keyring(monkeypatch):
    """Mock keyring to isolate tests from host keyring."""
    store: dict[tuple[str, str], str] = {}

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
def test_auth_manager_init_and_cache_loading(mock_keyring, mock_msal):
    """A keyring outage degrades to memory-only storage, never plaintext files."""
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"

    try:
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.client_id == "client_id"
        assert auth.secure_cache_available is False
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_load_token_cache_from_keyring(mock_keyring, mock_msal):
    mock_keyring[(auth_mod.SERVICE_NAME, auth_mod.TOKEN_CACHE_ACCOUNT)] = "cache_data"
    auth = AuthManager("client_id", "authority", ["User.Read"])
    auth.token_cache.deserialize.assert_called_with("cache_data")


@pytest.mark.concept("ECO-4.1")
def test_load_token_cache_keyring_error(mock_keyring, mock_msal):
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth = AuthManager("client_id", "authority", ["User.Read"])
        assert auth.secure_cache_available is False
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_save_token_cache(mock_keyring, mock_msal):
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

    # A keyring failure leaves the cache in memory and marks secure storage absent.
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth.save_token_cache()
        assert auth.secure_cache_available is False
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_load_selected_account(mock_keyring, mock_msal):
    key = (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY)
    mock_keyring[key] = "invalid json"
    auth = AuthManager("client_id", "authority", ["User.Read"])
    assert auth.selected_account_id is None

    mock_keyring[key] = '{"account_id": "test_account_id"}'
    auth = AuthManager("client_id", "authority", ["User.Read"])
    assert auth.selected_account_id == "test_account_id"


@pytest.mark.concept("ECO-4.1")
def test_save_selected_account(mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No selected account
    auth.selected_account_id = None
    auth.save_selected_account()
    assert (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY) not in mock_keyring

    # Case 2: Save to keyring
    auth.selected_account_id = "keyring_acc"
    auth.save_selected_account()
    assert "keyring_acc" in mock_keyring.get(
        (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY)
    )

    # Case 3: Keyring failure leaves the selection in memory only.
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth.selected_account_id = "memory_acc"
        auth.save_selected_account()
        assert auth.secure_cache_available is False
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
def test_get_current_account(mock_keyring, mock_msal):
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

    # Case 4: A stale selection fails closed rather than switching identities.
    auth.selected_account_id = "missing"
    assert auth.get_current_account() is None


@pytest.mark.concept("ECO-4.1")
def test_get_token(mock_keyring, mock_msal):
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
def test_get_token_details(mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"])

    # Case 1: No account
    mock_msal.get_accounts.return_value = []
    assert auth.get_token_details() is None

    # Case 2: Success for the configured scopes.
    acc = {"home_account_id": "acc1"}
    mock_msal.get_accounts.return_value = [acc]
    mock_msal.acquire_token_silent.return_value = {
        "access_token": "token123",
        "expires_in": 3600,
    }
    res = auth.get_token_details()
    assert res is not None
    assert res["access_token"] == "token123"
    mock_msal.acquire_token_silent.assert_called_with(
        auth.scopes,
        account=acc,
    )

    # Case 3: Silent return is empty
    mock_msal.acquire_token_silent.return_value = {}
    assert auth.get_token_details() is None


@pytest.mark.concept("ECO-4.1")
def test_acquire_token_by_device_code(mock_keyring, mock_msal):
    auth = AuthManager("client_id", "authority", ["User.Read"], allow_device_code=True)

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
    with pytest.raises(
        Exception,
        match="Authentication failed: Microsoft identity rejected the request",
    ):
        auth.acquire_token_by_device_code(callback)


@pytest.mark.concept("ECO-4.1")
def test_logout_and_account_management(mock_keyring, mock_msal):
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

    # Test logout with existing keyring entries.
    auth.token_cache.has_state_changed = True
    auth.token_cache.serialize.return_value = "cache_data"
    auth.selected_account_id = "some_id"
    auth.access_token = "some_token"

    auth.save_token_cache()
    auth.save_selected_account()
    assert (auth_mod.SERVICE_NAME, auth_mod.TOKEN_CACHE_ACCOUNT) in mock_keyring
    assert (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY) in mock_keyring

    auth.logout()
    assert auth.selected_account_id is None
    assert auth.access_token is None
    assert (auth_mod.SERVICE_NAME, auth_mod.TOKEN_CACHE_ACCOUNT) not in mock_keyring
    assert (auth_mod.SERVICE_NAME, auth_mod.SELECTED_ACCOUNT_KEY) not in mock_keyring

    # A keyring delete failure is contained during logout.
    original_service = auth_mod.SERVICE_NAME
    auth_mod.SERVICE_NAME = "fail"
    try:
        auth2 = AuthManager("client_id", "authority", ["User.Read"])
        auth2.logout()  # Should not raise exception
    finally:
        auth_mod.SERVICE_NAME = original_service


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_uses_authenticated_manager():
    manager = MagicMock()
    manager.get_token.return_value = "cached-token"
    with (
        patch("microsoft_agent.auth.get_auth_manager", return_value=manager),
        patch("microsoft_agent.api_client.MicrosoftGraphApi") as mock_api,
    ):
        client = await get_client()
    assert client is mock_api.return_value
    mock_api.assert_called_once_with(manager)


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_requires_cached_token():
    manager = MagicMock()
    manager.get_token.return_value = None
    with (
        patch("microsoft_agent.auth.get_auth_manager", return_value=manager),
        pytest.raises(ValueError, match="Microsoft token is not available"),
    ):
        await get_client()


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_propagates_configuration_error():
    with (
        patch(
            "microsoft_agent.auth.get_auth_manager",
            side_effect=ValueError("Microsoft authentication is not configured"),
        ),
        pytest.raises(ValueError, match="authentication is not configured"),
    ):
        await get_client()


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_wraps_manager_auth_error():
    with (
        patch(
            "microsoft_agent.auth.get_auth_manager",
            side_effect=AuthError("sensitive provider detail"),
        ),
        pytest.raises(RuntimeError, match="Microsoft credentials were rejected"),
    ):
        await get_client()


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_get_client_wraps_api_instantiation_error():
    manager = MagicMock()
    manager.get_token.return_value = "cached-token"
    with (
        patch("microsoft_agent.auth.get_auth_manager", return_value=manager),
        patch(
            "microsoft_agent.api_client.MicrosoftGraphApi",
            side_effect=UnauthorizedError("sensitive provider detail"),
        ),
        pytest.raises(RuntimeError, match="Microsoft credentials were rejected"),
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
