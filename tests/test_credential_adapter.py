"""Tests for credential_adapter.py module."""

import pytest
from azure.core.credentials import AccessToken

from microsoft_agent.credential_adapter import AuthManagerCredential


class TestAuthManagerCredential:
    """Test AuthManagerCredential class."""

    def test_init(self, mock_auth_manager):
        """Test AuthManagerCredential initialization."""
        credential = AuthManagerCredential(mock_auth_manager)
        assert credential.auth_manager == mock_auth_manager

    def test_get_token_success_with_expires_on(
        self, mock_auth_manager, mock_token_response
    ):
        """Test get_token successfully with expires_on."""
        mock_auth_manager.get_token_details.return_value = mock_token_response
        credential = AuthManagerCredential(mock_auth_manager)

        result = credential.get_token("https://graph.microsoft.com/.default")

        assert isinstance(result, AccessToken)
        assert result.token == "test_access_token"
        assert result.expires_on == 1234567890

    def test_get_token_success_without_expires_on(self, mock_auth_manager):
        """Test get_token successfully without expires_on."""
        import time

        mock_token_response = {"access_token": "test_access_token", "expires_in": 3600}
        mock_auth_manager.get_token_details.return_value = mock_token_response
        credential = AuthManagerCredential(mock_auth_manager)

        result = credential.get_token("https://graph.microsoft.com/.default")

        assert isinstance(result, AccessToken)
        assert result.token == "test_access_token"
        assert result.expires_on > int(time.time())

    def test_get_token_no_token_details(self, mock_auth_manager):
        """Test get_token when no token details available."""
        mock_auth_manager.get_token_details.return_value = None
        credential = AuthManagerCredential(mock_auth_manager)

        with pytest.raises(Exception, match="Failed to acquire token"):
            credential.get_token("https://graph.microsoft.com/.default")

    def test_get_token_no_access_token(self, mock_auth_manager):
        """Test get_token when token details has no access_token."""
        mock_auth_manager.get_token_details.return_value = {"other_field": "value"}
        credential = AuthManagerCredential(mock_auth_manager)

        with pytest.raises(Exception, match="Failed to acquire token"):
            credential.get_token("https://graph.microsoft.com/.default")

    def test_get_token_with_additional_params(
        self, mock_auth_manager, mock_token_response
    ):
        """Test get_token with additional parameters."""
        mock_auth_manager.get_token_details.return_value = mock_token_response
        credential = AuthManagerCredential(mock_auth_manager)

        result = credential.get_token(
            "https://graph.microsoft.com/.default",
            _claims="test_claims",
            _tenant_id="test_tenant",
            custom_param="custom_value",
        )

        assert isinstance(result, AccessToken)
        assert result.token == "test_access_token"

    def test_get_token_multiple_scopes(self, mock_auth_manager, mock_token_response):
        """Test get_token with multiple scopes."""
        mock_auth_manager.get_token_details.return_value = mock_token_response
        credential = AuthManagerCredential(mock_auth_manager)

        result = credential.get_token(
            "https://graph.microsoft.com/.default",
            "https://graph.microsoft.com/mail.read",
        )

        assert isinstance(result, AccessToken)
        assert result.token == "test_access_token"
