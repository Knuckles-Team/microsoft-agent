import time
from unittest.mock import MagicMock

import pytest

from microsoft_agent.credential_adapter import AuthManagerCredential


@pytest.mark.concept("ECO-4.1")
def test_credential_adapter_success():
    mock_auth = MagicMock()
    # 1. token details with expires_on
    mock_auth.get_token_details.return_value = {
        "access_token": "test_token",
        "expires_on": 123456789,
    }
    cred = AuthManagerCredential(mock_auth)
    token = cred.get_token("scope")
    assert token.token == "test_token"
    assert token.expires_on == 123456789


@pytest.mark.concept("ECO-4.1")
def test_credential_adapter_expires_in():
    mock_auth = MagicMock()
    # 2. token details without expires_on but with expires_in
    mock_auth.get_token_details.return_value = {
        "access_token": "test_token",
        "expires_in": 3600,
    }
    cred = AuthManagerCredential(mock_auth)
    token = cred.get_token("scope")
    assert token.token == "test_token"
    assert token.expires_on > time.time()


@pytest.mark.concept("ECO-4.1")
def test_credential_adapter_fail():
    mock_auth = MagicMock()
    # 3. no token_details or no access_token
    mock_auth.get_token_details.return_value = None
    cred = AuthManagerCredential(mock_auth)
    with pytest.raises(Exception, match="Failed to acquire token"):
        cred.get_token("scope")
