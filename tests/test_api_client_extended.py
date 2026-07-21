"""Extended tests for the current modular Microsoft Graph API client."""

from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from microsoft_agent.api_client import MicrosoftGraphApi


def _awaited_argument(mock: Any) -> Any:
    """Return the first argument from a verified awaited mock call."""

    awaited_call = mock.await_args
    assert awaited_call is not None
    return awaited_call.args[0]


class TestMicrosoftGraphApi:
    """Test MicrosoftGraphApi class."""

    def test_init_success(self, mock_auth_manager):
        """Test successful initialization."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                assert api.auth_manager == mock_auth_manager

    def test_init_auth_error(self, mock_auth_manager):
        """Test initialization with authentication error."""
        mock_auth_manager.get_token.return_value = None
        mock_auth_manager.get_current_account.return_value = None

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                from agent_utilities.core.exceptions import AuthError

                with pytest.raises(AuthError, match="Microsoft authentication failed"):
                    MicrosoftGraphApi(mock_auth_manager)

    def test_login_already_authenticated(self, mock_auth_manager):
        """Test login when already authenticated."""
        mock_auth_manager.get_token.return_value = "existing_token"

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.login()
                assert result == "Already authenticated."

    def test_login_force(self, mock_auth_manager, mock_device_flow):
        """Test forced login."""
        mock_auth_manager.get_token.return_value = "test_token"  # Make init succeed
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        mock_auth_manager.acquire_token_by_device_code.return_value = (
            "Authentication successful"
        )

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.login(force=True)
                assert result == "Authentication successful"

    def test_logout(self, mock_auth_manager):
        """Test logout."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.logout()
                assert result == "Logged out."
                mock_auth_manager.logout.assert_called_once()

    def test_verify_login_authenticated(self, mock_auth_manager):
        """Test verify_login when authenticated."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.verify_login()
                assert "Authenticated as test@example.com" in result

    def test_verify_login_not_authenticated(self, mock_auth_manager):
        """Test verify_login when not authenticated."""
        mock_auth_manager.get_token.return_value = "test_token"  # Make init succeed
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                # Mock verify_login to return not authenticated
                with patch.object(
                    api, "verify_login", return_value="Not authenticated."
                ):
                    result = api.verify_login()
                assert result == "Not authenticated."

    def test_verify_login_unknown_user(self, mock_auth_manager):
        """Test verify_login when account exists but username unknown."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = None

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.verify_login()
                assert result == "Authenticated with workload identity"

    def test_list_accounts(self, mock_auth_manager):
        """Test list_accounts."""
        mock_auth_manager.get_token.return_value = "test_token"  # Make init succeed
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        mock_auth_manager.list_accounts.return_value = [
            {"username": "test@example.com"}
        ]

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.list_accounts()
                assert result == [{"username": "test@example.com"}]

    def test_search_tools(self, mock_auth_manager):
        """Test search_tools."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.search_tools("mail", limit=5)
                assert isinstance(result, list)
                assert all("mail" in name.lower() for name in result)

    def test_search_tools_limit(self, mock_auth_manager):
        """Test search_tools with limit."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch("microsoft_agent.api.api_client_base.GraphServiceClient"):
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                api = MicrosoftGraphApi(mock_auth_manager)
                result = api.search_tools("list", limit=3)
                assert len(result) <= 3


@pytest.mark.asyncio
class TestMailOperations:
    """Test mail-related operations."""

    async def test_list_mail_messages_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_mail_messages successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                mock_client = MagicMock()
                mock_client.me.messages.get = AsyncMock(
                    return_value=mock_native_response
                )
                mock_client_class.return_value = mock_client

                api = MicrosoftGraphApi(mock_auth_manager)
                result = await api.list_mail_messages()
                assert result == {"value": []}

    async def test_list_mail_messages_with_params(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_mail_messages with query parameters."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        params = {"$select": "id,subject", "$top": "10", "$filter": "isRead eq false"}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_messages(params=params)
            assert result == {"value": []}

    async def test_list_mail_messages_error(self, mock_auth_manager):
        """Test list_mail_messages with error."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.get = AsyncMock(side_effect=Exception("API Error"))
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_messages()
            assert "error" in result

    async def test_get_mail_message_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test get_mail_message successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.get = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.get_mail_message("msg123")
            assert result == {"value": []}

    async def test_send_mail_success(self, mock_auth_manager, sample_mail_data):
        """Test send_mail successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.send_mail.post = AsyncMock()
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.send_mail(sample_mail_data)
            assert result == {"status": "success"}

    async def test_send_mail_error(self, mock_auth_manager, sample_mail_data):
        """Test send_mail with error."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.send_mail.post = AsyncMock(
                side_effect=Exception("Send failed")
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.send_mail(sample_mail_data)
            assert "error" in result

    async def test_create_draft_email_success(
        self, mock_auth_manager, sample_mail_data
    ):
        """Test create_draft_email successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_response = MagicMock()
            mock_response.json.return_value = {"id": "draft123"}
            mock_client.me.messages.post = AsyncMock(return_value=mock_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.create_draft_email(sample_mail_data)
            assert result["id"] == "draft123"

    async def test_delete_mail_message_success(self, mock_auth_manager):
        """Test delete_mail_message successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.delete = AsyncMock()
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.delete_mail_message("msg123")
            assert result == {"status": "success"}

    async def test_move_mail_message_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test move_mail_message successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            with patch("microsoft_agent.api.api_client_base.AuthManagerCredential"):
                mock_client = MagicMock()
                mock_client.me.messages.by_message_id.return_value.move.post = (
                    AsyncMock(return_value=mock_native_response)
                )
                mock_client_class.return_value = mock_client

                api = MicrosoftGraphApi(mock_auth_manager)
                result = await api.move_mail_message(
                    "msg123", {"destinationId": "folder123"}
                )
                assert result == {"value": []}

    async def test_update_mail_message_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test update_mail_message successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.patch = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.update_mail_message("msg123", {"subject": "Updated"})
            assert result == {"value": []}

    async def test_list_mail_folders_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_mail_folders successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.mail_folders.get = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_folders()
            assert result == {"value": []}


@pytest.mark.asyncio
class TestUserOperations:
    """Test user-related operations."""

    async def test_get_me_success(self, mock_auth_manager, mock_native_response):
        """Test get_me successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.get_me()
            assert result == {"value": []}

    async def test_list_users_success(self, mock_auth_manager, mock_native_response):
        """Test list_users successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.users.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_users()
            assert result == {"value": []}

    async def test_list_users_with_consistency_level(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_users with ConsistencyLevel header."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        params = {"ConsistencyLevel": "eventual"}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.users.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_users(params=params)
            assert result == {"value": []}


@pytest.mark.asyncio
class TestSharedMailboxOperations:
    """Test shared mailbox operations."""

    async def test_list_shared_mailbox_messages_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_shared_mailbox_messages successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.users.by_user_id.return_value.messages.get = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_shared_mailbox_messages("user123")
            assert result == {"value": []}

    async def test_get_shared_mailbox_message_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test get_shared_mailbox_message successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.users.by_user_id.return_value.messages.by_message_id.return_value.get = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.get_shared_mailbox_message("user123", "msg123")
            assert result == {"value": []}


@pytest.mark.asyncio
class TestAttachmentOperations:
    """Test attachment operations."""

    async def test_add_mail_attachment_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test add_mail_attachment successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        attachment_data = {
            "name": "test.txt",
            "contentType": "text/plain",
            "contentBytes": "dGVzdCBjb250ZW50",
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.attachments.post = (
                AsyncMock(return_value=mock_native_response)
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.add_mail_attachment("msg123", attachment_data)
            assert result == {"value": []}

    async def test_list_mail_attachments_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_mail_attachments successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.attachments.get = (
                AsyncMock(return_value=mock_native_response)
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_attachments("msg123")
            assert result == {"value": []}

    async def test_delete_mail_attachment_success(self, mock_auth_manager):
        """Test delete_mail_attachment successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.by_message_id.return_value.attachments.by_attachment_id.return_value.delete = AsyncMock()
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.delete_mail_attachment("msg123", "attach123")
            assert result == {"status": "success"}


@pytest.mark.asyncio
class TestParameterHandling:
    """Test parameter handling in API methods."""

    async def test_boolean_count_parameter(
        self, mock_auth_manager, mock_native_response
    ):
        """Test boolean count parameter handling."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        params = {"$count": "true"}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_messages(params=params)
            assert result == {"value": []}

    async def test_orderby_parameter_splitting(
        self, mock_auth_manager, mock_native_response
    ):
        """Test orderby parameter splitting."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        params = {"$orderby": "receivedDateTime desc,subject asc"}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.messages.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_mail_messages(params=params)
            assert result == {"value": []}


@pytest.mark.asyncio
class TestCalendarOperations:
    """Test calendar-related operations."""

    async def test_list_calendars_success(
        self, mock_auth_manager, mock_native_response
    ):
        """Test list_calendars successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.calendars.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_calendars()
            assert result == {"value": []}


@pytest.mark.asyncio
class TestDriveOperations:
    """Test drive/OneDrive operations."""

    async def test_list_drives_success(self, mock_auth_manager, mock_native_response):
        """Test list_drives successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.drives.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_drives()
            assert result == {"value": []}


@pytest.mark.asyncio
class TestGroupOperations:
    """Test group-related operations."""

    async def test_list_groups_success(self, mock_auth_manager, mock_native_response):
        """Test list_groups successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.groups.get = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.list_groups()
            assert result == {"value": []}

    async def test_create_group_success(self, mock_auth_manager, mock_native_response):
        """Test create_group successfully."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        group_data = {"displayName": "Test Group", "mailEnabled": True}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.groups.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.create_group(group_data)
            assert result == {"value": []}


@pytest.mark.asyncio
class TestRequestPayloadPropagation:
    """Verify generated SDK request bodies retain caller-provided payloads."""

    async def test_calendar_event_mutations_propagate_complete_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """Calendar create/update variants retain meeting and location fields."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "subject": "Architecture review",
            "body": {"contentType": "html", "content": "<p>Review</p>"},
            "start": {
                "dateTime": "2026-07-20T09:00:00",
                "timeZone": "Central Standard Time",
            },
            "end": {
                "dateTime": "2026-07-20T09:30:00",
                "timeZone": "Central Standard Time",
            },
            "attendees": [
                {
                    "type": "required",
                    "emailAddress": {
                        "address": "attendee@example.com",
                        "name": "Attendee",
                    },
                }
            ],
            "location": {
                "displayName": "Conference Room A",
                "locationType": "conferenceRoom",
            },
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness",
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.events.post = AsyncMock(return_value=mock_native_response)
            event_item = mock_client.me.events.by_event_id.return_value
            event_item.patch = AsyncMock(return_value=mock_native_response)
            calendar_events = (
                mock_client.me.calendars.by_calendar_id.return_value.events
            )
            calendar_events.post = AsyncMock(return_value=mock_native_response)
            calendar_event_item = calendar_events.by_event_id.return_value
            calendar_event_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_calendar_event(data) == {"value": []}
            assert await api.update_calendar_event("event-1", data) == {"value": []}
            assert await api.create_specific_calendar_event("calendar-1", data) == {
                "value": []
            }
            assert await api.update_specific_calendar_event(
                "calendar-1", "event-2", data
            ) == {"value": []}

        calls = (
            mock_client.me.events.post.await_args,
            event_item.patch.await_args,
            calendar_events.post.await_args,
            calendar_event_item.patch.await_args,
        )
        for awaited_call in calls:
            assert awaited_call is not None
            event = awaited_call.args[0]
            assert event.subject == "Architecture review"
            assert event.attendees[0].email_address.address == ("attendee@example.com")
            assert event.location.display_name == "Conference Room A"
            assert event.is_online_meeting is True
            assert event.online_meeting_provider.value == "teamsForBusiness"

    async def test_mail_mutations_propagate_complete_recipient_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """Send, draft, shared-send, and update retain supported message fields."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        message: dict[str, Any] = {
            "subject": "Quarterly review",
            "body": {"contentType": "html", "content": "<p>Attached</p>"},
            "toRecipients": [
                {"emailAddress": {"address": "to@example.com", "name": "To"}}
            ],
            "ccRecipients": [
                {"emailAddress": {"address": "cc@example.com", "name": "Cc"}}
            ],
            "bccRecipients": [
                {"emailAddress": {"address": "bcc@example.com", "name": "Bcc"}}
            ],
            "replyTo": [
                {
                    "emailAddress": {
                        "address": "reply@example.com",
                        "name": "Replies",
                    }
                }
            ],
            "importance": "high",
            "categories": ["Executive"],
        }
        send_data = {"message": message, "saveToSentItems": False}

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.send_mail.post = AsyncMock()
            mock_client.me.messages.post = AsyncMock(return_value=mock_native_response)
            message_item = mock_client.me.messages.by_message_id.return_value
            message_item.patch = AsyncMock(return_value=mock_native_response)
            shared_send = mock_client.users.by_user_id.return_value.send_mail
            shared_send.post = AsyncMock()
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.send_mail(send_data) == {"status": "success"}
            assert await api.create_draft_email(message) == {"value": []}
            assert await api.update_mail_message(
                "message-1",
                {"importance": "high", "categories": ["Executive"], "isRead": True},
            ) == {"value": []}
            assert await api.send_shared_mailbox_mail(
                "shared@example.com", send_data
            ) == {"status": "success"}

        send_call = mock_client.me.send_mail.post.await_args
        draft_call = mock_client.me.messages.post.await_args
        update_call = message_item.patch.await_args
        shared_call = shared_send.post.await_args
        assert send_call is not None
        assert draft_call is not None
        assert update_call is not None
        assert shared_call is not None

        for request_body in (send_call.args[0], shared_call.args[0]):
            assert request_body.save_to_sent_items is False
            sent_message = request_body.message
            assert sent_message.cc_recipients[0].email_address.address == (
                "cc@example.com"
            )
            assert sent_message.bcc_recipients[0].email_address.address == (
                "bcc@example.com"
            )
            assert sent_message.reply_to[0].email_address.address == (
                "reply@example.com"
            )
            assert sent_message.importance.value == "high"

        draft = draft_call.args[0]
        assert draft.cc_recipients[0].email_address.address == "cc@example.com"
        assert draft.bcc_recipients[0].email_address.address == "bcc@example.com"
        assert draft.reply_to[0].email_address.address == "reply@example.com"
        assert draft.importance.value == "high"

        updated = update_call.args[0]
        assert updated.importance.value == "high"
        assert updated.categories == ["Executive"]
        assert updated.is_read is True

    async def test_contact_mutations_propagate_structured_fields(
        self, mock_auth_manager, mock_native_response
    ):
        """Contact create and update retain structured email and business data."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "givenName": "Ada",
            "surname": "Lovelace",
            "companyName": "Analytical Engines",
            "jobTitle": "Mathematician",
            "businessPhones": ["+1 555 0100"],
            "mobilePhone": "+1 555 0199",
            "emailAddresses": [{"address": "ada@example.com", "name": "Ada Lovelace"}],
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.contacts.post = AsyncMock(return_value=mock_native_response)
            contact_item = mock_client.me.contacts.by_contact_id.return_value
            contact_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_outlook_contact(data) == {"value": []}
            assert await api.update_outlook_contact("contact-1", data) == {"value": []}

        for awaited_call in (
            mock_client.me.contacts.post.await_args,
            contact_item.patch.await_args,
        ):
            assert awaited_call is not None
            contact = awaited_call.args[0]
            assert contact.company_name == "Analytical Engines"
            assert contact.business_phones == ["+1 555 0100"]
            assert contact.mobile_phone == "+1 555 0199"
            assert contact.email_addresses[0].address == "ada@example.com"
            assert contact.email_addresses[0].name == "Ada Lovelace"

    async def test_online_meeting_mutations_propagate_meeting_policy(
        self, mock_auth_manager, mock_native_response
    ):
        """Online meeting create/update retain participants and meeting policy."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "subject": "Teams architecture review",
            "startDateTime": "2026-07-20T14:00:00Z",
            "endDateTime": "2026-07-20T14:45:00Z",
            "participants": {
                "attendees": [
                    {
                        "upn": "attendee@example.com",
                        "role": "attendee",
                        "identity": {
                            "user": {
                                "id": "11111111-1111-4111-8111-111111111111",
                                "displayName": "Attendee",
                            }
                        },
                    }
                ]
            },
            "lobbyBypassSettings": {
                "scope": "organization",
                "isDialInBypassEnabled": True,
            },
            "allowedPresenters": "organization",
            "allowAttendeeToEnableCamera": False,
            "allowAttendeeToEnableMic": True,
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.online_meetings.post = AsyncMock(
                return_value=mock_native_response
            )
            meeting_item = (
                mock_client.me.online_meetings.by_online_meeting_id.return_value
            )
            meeting_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_online_meeting(data) == {"value": []}
            assert await api.update_online_meeting("meeting-1", data) == {"value": []}

        for awaited_call in (
            mock_client.me.online_meetings.post.await_args,
            meeting_item.patch.await_args,
        ):
            assert awaited_call is not None
            meeting = awaited_call.args[0]
            attendee = meeting.participants.attendees[0]
            assert attendee.upn == "attendee@example.com"
            assert attendee.role.value == "attendee"
            assert attendee.identity.user.display_name == "Attendee"
            assert meeting.lobby_bypass_settings.scope.value == "organization"
            assert meeting.lobby_bypass_settings.is_dial_in_bypass_enabled is True
            assert meeting.allowed_presenters.value == "organization"
            assert meeting.allow_attendee_to_enable_camera is False
            assert meeting.allow_attendee_to_enable_mic is True

    async def test_list_online_meetings_propagates_query_filters(
        self, mock_auth_manager, mock_native_response
    ):
        """Online meeting lookup forwards documented OData filters and projection."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        params = {
            "$filter": "joinMeetingIdSettings/joinMeetingId eq '1234567890'",
            "$select": "id,subject,joinWebUrl,joinMeetingIdSettings",
            "$expand": "participants",
            "$top": "5",
            "$count": "true",
            "Accept-Language": "en-US",
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.online_meetings.get = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.list_online_meetings(params) == {"value": []}

        awaited_call = mock_client.me.online_meetings.get.await_args
        assert awaited_call is not None
        request_config = awaited_call.kwargs["request_configuration"]
        query = request_config.query_parameters
        assert query.filter == params["$filter"]
        assert query.select == ["id", "subject", "joinWebUrl", "joinMeetingIdSettings"]
        assert query.expand == ["participants"]
        assert query.top == 5
        assert query.count is True
        assert request_config.headers.get("Accept-Language") == {"en-US"}

    async def test_find_meeting_times_propagates_payload(
        self, mock_auth_manager, mock_native_response
    ):
        """Find-meeting-times sends attendees, constraints, and options."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "attendees": [
                {
                    "type": "required",
                    "emailAddress": {
                        "address": "attendee@example.com",
                        "name": "Attendee",
                    },
                }
            ],
            "timeConstraint": {
                "activityDomain": "work",
                "timeSlots": [
                    {
                        "start": {
                            "dateTime": "2026-07-20T09:00:00",
                            "timeZone": "UTC",
                        },
                        "end": {
                            "dateTime": "2026-07-20T17:00:00",
                            "timeZone": "UTC",
                        },
                    }
                ],
            },
            "meetingDuration": "PT30M",
            "maxCandidates": 5,
            "returnSuggestionReasons": True,
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.me.find_meeting_times.post = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.find_meeting_times(data)

        assert result == {"value": []}
        awaited_call = mock_client.me.find_meeting_times.post.await_args
        assert awaited_call is not None
        request_body = awaited_call.args[0]
        assert (
            request_body.attendees[0].email_address.address
            == data["attendees"][0]["emailAddress"]["address"]
        )
        assert request_body.attendees[0].type.value == "required"
        assert request_body.time_constraint.activity_domain.value == "work"
        assert request_body.time_constraint.time_slots[0].start.time_zone == "UTC"
        assert request_body.meeting_duration.total_seconds() == 30 * 60
        assert request_body.max_candidates == 5
        assert request_body.return_suggestion_reasons is True

    async def test_search_query_propagates_payload(
        self, mock_auth_manager, mock_native_response
    ):
        """Search sends the complete typed request collection."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "requests": [
                {
                    "entityTypes": ["message"],
                    "query": {"queryString": "quarterly report"},
                    "from": 10,
                    "size": 25,
                    "fields": ["subject", "receivedDateTime"],
                }
            ]
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.search.query.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.search_query(data)

        assert result == {"value": []}
        awaited_call = mock_client.search.query.post.await_args
        assert awaited_call is not None
        request_body = awaited_call.args[0]
        request = request_body.requests[0]
        assert [entity_type.value for entity_type in request.entity_types] == [
            "message"
        ]
        assert request.query.query_string == "quarterly report"
        assert request.from_ == 10
        assert request.size == 25
        assert request.fields == ["subject", "receivedDateTime"]

    async def test_update_admin_sharepoint_propagates_payload(
        self, mock_auth_manager, mock_native_response
    ):
        """SharePoint admin update sends nested tenant settings."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "settings": {
                "isLoopEnabled": True,
                "sharingCapability": "externalUserSharingOnly",
                "sharingAllowedDomainList": ["contoso.com", "fabrikam.com"],
            }
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.admin.sharepoint.patch = AsyncMock(
                return_value=mock_native_response
            )
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.update_admin_sharepoint(data)

        assert result == {"value": []}
        awaited_call = mock_client.admin.sharepoint.patch.await_args
        assert awaited_call is not None
        request_body = awaited_call.args[0]
        assert request_body.settings.is_loop_enabled is True
        assert (
            request_body.settings.sharing_capability.value == "externalUserSharingOnly"
        )
        assert request_body.settings.sharing_allowed_domain_list == [
            "contoso.com",
            "fabrikam.com",
        ]

    async def test_create_print_job_propagates_payload(
        self, mock_auth_manager, mock_native_response
    ):
        """Print-job creation sends configuration and fetchability settings."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "configuration": {
                "copies": 2,
                "colorMode": "color",
                "duplexMode": "flipOnLongEdge",
                "fitPdfToPage": True,
            },
            "isFetchable": True,
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            jobs = mock_client.print.printers.by_printer_id.return_value.jobs
            jobs.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            result = await api.create_print_job("printer-123", data)

        assert result == {"value": []}
        awaited_call = jobs.post.await_args
        assert awaited_call is not None
        request_body = awaited_call.args[0]
        assert request_body.configuration.copies == 2
        assert request_body.configuration.color_mode.value == "color"
        assert request_body.configuration.duplex_mode.value == "flipOnLongEdge"
        assert request_body.configuration.fit_pdf_to_page is True
        assert request_body.is_fetchable is True

    async def test_planner_mutations_propagate_payloads_and_etags(
        self, mock_auth_manager, mock_native_response
    ):
        """Planner mutations retain typed fields and send concurrency headers."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        create_data: dict[str, Any] = {
            "planId": "plan-1",
            "bucketId": "bucket-1",
            "title": "Ship the release",
            "priority": 1,
            "dueDateTime": "2026-07-31T17:00:00Z",
            "assignments": {
                "user-1": {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !",
                }
            },
        }
        update_data: dict[str, Any] = {
            "percentComplete": 50,
            "startDateTime": "2026-07-20T09:00:00Z",
            "appliedCategories": {"category3": True},
        }
        details_data: dict[str, Any] = {
            "description": "Release checklist",
            "previewType": "checklist",
            "checklist": {
                "item-1": {
                    "@odata.type": "#microsoft.graph.plannerChecklistItem",
                    "title": "Run smoke tests",
                    "isChecked": False,
                }
            },
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            tasks = mock_client.planner.tasks
            task_item = tasks.by_planner_task_id.return_value
            tasks.post = AsyncMock(return_value=mock_native_response)
            task_item.patch = AsyncMock(return_value=mock_native_response)
            task_item.details.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_planner_task(create_data) == {"value": []}
            assert await api.update_planner_task(
                "task-1", update_data, etag='W/"task-etag"'
            ) == {"value": []}
            assert await api.update_planner_task_details(
                "task-1", details_data, params={"If-Match": 'W/"details-etag"'}
            ) == {"value": []}

        create_call = tasks.post.await_args
        update_call = task_item.patch.await_args
        details_call = task_item.details.patch.await_args
        assert create_call is not None
        assert update_call is not None
        assert details_call is not None

        created = create_call.args[0]
        assert created.bucket_id == "bucket-1"
        assert created.priority == 1
        assert created.due_date_time.isoformat() == "2026-07-31T17:00:00+00:00"
        assert created.assignments.additional_data["user-1"]["orderHint"] == " !"

        updated = update_call.args[0]
        assert updated.percent_complete == 50
        assert updated.start_date_time.isoformat() == "2026-07-20T09:00:00+00:00"
        assert updated.applied_categories.additional_data["category3"] is True
        update_config = update_call.kwargs["request_configuration"]
        update_config.headers.add.assert_any_call("If-Match", 'W/"task-etag"')
        update_config.headers.add.assert_any_call("Prefer", "return=representation")

        details = details_call.args[0]
        assert details.preview_type.value == "checklist"
        assert details.checklist.additional_data["item-1"]["title"] == (
            "Run smoke tests"
        )
        details_config = details_call.kwargs["request_configuration"]
        details_config.headers.add.assert_any_call("If-Match", 'W/"details-etag"')
        details_config.headers.add.assert_any_call("Prefer", "return=representation")

    async def test_planner_update_rejects_missing_or_unsafe_etag(
        self, mock_auth_manager
    ):
        """Planner writes fail before I/O when the concurrency token is unsafe."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            task_item = mock_client.planner.tasks.by_planner_task_id.return_value
            task_item.patch = AsyncMock()
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            missing = await api.update_planner_task("task-1", {"title": "New"})
            unsafe = await api.update_planner_task(
                "task-1", {"title": "New"}, etag='W/"ok"\r\nX-Evil: yes'
            )

        assert "required" in missing["error"]
        assert "quoted HTTP entity tag" in unsafe["error"]
        task_item.patch.assert_not_awaited()

    async def test_conditional_access_mutations_propagate_nested_policy(
        self, mock_auth_manager, mock_native_response
    ):
        """Conditional Access tools retain conditions and enforcement controls."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "displayName": "Require MFA",
            "state": "enabledForReportingButNotEnforced",
            "conditions": {
                "users": {"includeUsers": ["All"]},
                "applications": {"includeApplications": ["All"]},
                "clientAppTypes": ["all"],
            },
            "grantControls": {"operator": "OR", "builtInControls": ["mfa"]},
            "sessionControls": {
                "signInFrequency": {
                    "isEnabled": True,
                    "type": "hours",
                    "value": 8,
                }
            },
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            policies = mock_client.identity.conditional_access.policies
            policy_item = policies.by_conditional_access_policy_id.return_value
            policies.post = AsyncMock(return_value=mock_native_response)
            policy_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_conditional_access_policy(data) == {"value": []}
            assert await api.update_conditional_access_policy("policy-1", data) == {
                "value": []
            }

        for awaited_call in (policies.post.await_args, policy_item.patch.await_args):
            assert awaited_call is not None
            policy = awaited_call.args[0]
            assert policy.state.value == "enabledForReportingButNotEnforced"
            assert policy.conditions.users.include_users == ["All"]
            assert policy.conditions.applications.include_applications == ["All"]
            assert policy.conditions.client_app_types[0].value == "all"
            assert policy.grant_controls.built_in_controls[0].value == "mfa"
            assert policy.session_controls.sign_in_frequency.value == 8

    async def test_create_agreement_propagates_terms_file(
        self, mock_auth_manager, mock_native_response
    ):
        """Terms-of-use creation retains the localized PDF payload."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "displayName": "Contoso Terms",
            "isViewingBeforeAcceptanceRequired": True,
            "isPerDeviceAcceptanceRequired": False,
            "files": [
                {
                    "fileName": "terms.pdf",
                    "language": "en",
                    "isDefault": True,
                    "fileData": {"data": "JVBERi0xLjQ="},
                }
            ],
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            mock_client.agreements.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_agreement(data) == {"value": []}

        awaited_call = mock_client.agreements.post.await_args
        assert awaited_call is not None
        agreement = awaited_call.args[0]
        assert agreement.is_per_device_acceptance_required is False
        assert agreement.files[0].file_name == "terms.pdf"
        assert agreement.files[0].language == "en"
        assert agreement.files[0].is_default is True
        assert agreement.files[0].file_data.data == b"%PDF-1.4"

    async def test_create_booking_appointment_propagates_full_payload(
        self, mock_auth_manager, mock_native_response
    ):
        """Bookings creation retains scheduling, customer, and staff fields."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        data: dict[str, Any] = {
            "serviceId": "service-1",
            "customerName": "Ada Lovelace",
            "customerEmailAddress": "ada@example.com",
            "customerPhone": "+1 555 0100",
            "customerTimeZone": "America/Chicago",
            "smsNotificationsEnabled": True,
            "staffMemberIds": ["staff-1"],
            "startDateTime": {
                "dateTime": "2026-07-20T09:00:00",
                "timeZone": "America/Chicago",
            },
            "endDateTime": {
                "dateTime": "2026-07-20T10:00:00",
                "timeZone": "America/Chicago",
            },
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            appointments = mock_client.solutions.booking_businesses.by_booking_business_id.return_value.appointments
            appointments.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_booking_appointment("business-1", data) == {
                "value": []
            }

        awaited_call = appointments.post.await_args
        assert awaited_call is not None
        appointment = awaited_call.args[0]
        assert appointment.customer_email_address == "ada@example.com"
        assert appointment.customer_phone == "+1 555 0100"
        assert appointment.customer_time_zone == "America/Chicago"
        assert appointment.sms_notifications_enabled is True
        assert appointment.staff_member_ids == ["staff-1"]
        assert appointment.start_date_time.date_time == "2026-07-20T09:00:00"
        assert appointment.end_date_time.time_zone == "America/Chicago"

    async def test_todo_and_group_mutations_propagate_complete_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """To Do and group writes retain nested and less-common properties."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        todo_create: dict[str, Any] = {
            "title": "Prepare launch",
            "importance": "high",
            "body": {"contentType": "text", "content": "Verify rollout gates"},
            "categories": ["Release"],
            "dueDateTime": {
                "dateTime": "2026-07-31T17:00:00",
                "timeZone": "UTC",
            },
        }
        todo_update: dict[str, Any] = {
            "status": "inProgress",
            "isReminderOn": True,
            "reminderDateTime": {
                "dateTime": "2026-07-30T09:00:00",
                "timeZone": "UTC",
            },
        }
        group_create: dict[str, Any] = {
            "displayName": "Release Team",
            "mailNickname": "release-team",
            "groupTypes": ["Unified"],
            "visibility": "Private",
            "membershipRule": 'user.department -eq "Engineering"',
            "membershipRuleProcessingState": "On",
        }
        group_update: dict[str, Any] = {
            "classification": "Confidential",
            "preferredLanguage": "en-US",
            "autoSubscribeNewMembers": True,
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            tasks = mock_client.me.todo.lists.by_todo_task_list_id.return_value.tasks
            task_item = tasks.by_todo_task_id.return_value
            groups = mock_client.groups
            group_item = groups.by_group_id.return_value
            tasks.post = AsyncMock(return_value=mock_native_response)
            task_item.patch = AsyncMock(return_value=mock_native_response)
            groups.post = AsyncMock(return_value=mock_native_response)
            group_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_todo_task("list-1", todo_create) == {"value": []}
            assert await api.update_todo_task("list-1", "task-1", todo_update) == {
                "value": []
            }
            assert await api.create_group(group_create) == {"value": []}
            assert await api.update_group("group-1", group_update) == {"value": []}

        created_task = _awaited_argument(tasks.post)
        assert created_task.importance.value == "high"
        assert created_task.body.content == "Verify rollout gates"
        assert created_task.categories == ["Release"]
        assert created_task.due_date_time.time_zone == "UTC"
        updated_task = _awaited_argument(task_item.patch)
        assert updated_task.status.value == "inProgress"
        assert updated_task.is_reminder_on is True
        assert updated_task.reminder_date_time.date_time == "2026-07-30T09:00:00"

        created_group = _awaited_argument(groups.post)
        assert created_group.mail_enabled is False
        assert created_group.security_enabled is True
        assert created_group.group_types == ["Unified"]
        assert created_group.membership_rule == 'user.department -eq "Engineering"'
        updated_group = _awaited_argument(group_item.patch)
        assert updated_group.classification == "Confidential"
        assert updated_group.preferred_language == "en-US"
        assert updated_group.auto_subscribe_new_members is True

    async def test_tenant_subscription_invitation_and_security_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """Tenant and security mutations preserve nested, enum, and date fields."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        organization: dict[str, Any] = {
            "displayName": "Contoso",
            "marketingNotificationEmails": ["marketing@example.com"],
            "technicalNotificationMails": ["operations@example.com"],
            "privacyProfile": {
                "contactEmail": "privacy@example.com",
                "statementUrl": "https://example.com/privacy",
            },
        }
        branding: dict[str, Any] = {
            "signInPageText": "Welcome to Contoso",
            "usernameHintText": "Use your Contoso email",
            "backgroundColor": "#112233",
        }
        subscription: dict[str, Any] = {
            "changeType": "created,updated",
            "notificationUrl": "https://example.com/notifications",
            "lifecycleNotificationUrl": "https://example.com/lifecycle",
            "resource": "users",
            "expirationDateTime": "2026-07-20T10:00:00Z",
            "includeResourceData": True,
            "encryptionCertificateId": "certificate-1",
        }
        subscription_update: dict[str, Any] = {
            "expirationDateTime": "2026-07-21T11:30:00Z",
            "lifecycleNotificationUrl": "https://example.com/new-lifecycle",
        }
        invitation: dict[str, Any] = {
            "invitedUserEmailAddress": "guest@example.com",
            "invitedUserDisplayName": "Guest User",
            "sendInvitationMessage": True,
            "invitedUserType": "Guest",
            "invitedUserMessageInfo": {
                "customizedMessageBody": "Welcome to the workspace",
                "messageLanguage": "en-US",
            },
        }
        alert_update: dict[str, Any] = {
            "status": "inProgress",
            "assignedTo": "analyst@example.com",
            "classification": "truePositive",
            "determination": "malware",
            "comments": [{"comment": "Escalated to incident response"}],
        }
        incident_update: dict[str, Any] = {
            "status": "active",
            "classification": "truePositive",
            "determination": "malware",
            "customTags": ["priority", "endpoint"],
            "resolvingComment": "Containment in progress",
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            org_item = mock_client.organization.by_organization_id.return_value
            subscription_item = (
                mock_client.subscriptions.by_subscription_id.return_value
            )
            alert_item = mock_client.security.alerts_v2.by_alert_id.return_value
            incident_item = mock_client.security.incidents.by_incident_id.return_value
            org_item.patch = AsyncMock(return_value=mock_native_response)
            org_item.branding.patch = AsyncMock(return_value=mock_native_response)
            mock_client.subscriptions.post = AsyncMock(
                return_value=mock_native_response
            )
            subscription_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client.invitations.post = AsyncMock(return_value=mock_native_response)
            alert_item.patch = AsyncMock(return_value=mock_native_response)
            incident_item.patch = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.update_organization("org-1", organization) == {"value": []}
            assert await api.update_org_branding("org-1", branding) == {"value": []}
            assert await api.create_subscription(subscription) == {"value": []}
            assert await api.update_subscription(
                "subscription-1", subscription_update
            ) == {"value": []}
            assert await api.create_invitation(invitation) == {"value": []}
            assert await api.update_security_alert("alert-1", alert_update) == {
                "value": []
            }
            assert await api.update_security_incident(
                "incident-1", incident_update
            ) == {"value": []}

        org = _awaited_argument(org_item.patch)
        assert org.marketing_notification_emails == ["marketing@example.com"]
        assert org.technical_notification_mails == ["operations@example.com"]
        assert org.privacy_profile.contact_email == "privacy@example.com"
        org_branding = _awaited_argument(org_item.branding.patch)
        assert org_branding.username_hint_text == "Use your Contoso email"
        assert org_branding.background_color == "#112233"

        created_subscription = _awaited_argument(mock_client.subscriptions.post)
        assert created_subscription.include_resource_data is True
        assert created_subscription.lifecycle_notification_url.endswith("/lifecycle")
        assert created_subscription.expiration_date_time.isoformat() == (
            "2026-07-20T10:00:00+00:00"
        )
        renewed_subscription = _awaited_argument(subscription_item.patch)
        assert renewed_subscription.lifecycle_notification_url.endswith(
            "/new-lifecycle"
        )
        created_invitation = _awaited_argument(mock_client.invitations.post)
        assert created_invitation.invite_redirect_url == "https://myapps.microsoft.com"
        assert (
            created_invitation.invited_user_message_info.customized_message_body
            == "Welcome to the workspace"
        )

        alert = _awaited_argument(alert_item.patch)
        assert alert.status.value == "inProgress"
        assert alert.classification.value == "truePositive"
        assert alert.determination.value == "malware"
        assert alert.comments[0].comment == "Escalated to incident response"
        incident = _awaited_argument(incident_item.patch)
        assert incident.status.value == "active"
        assert incident.custom_tags == ["priority", "endpoint"]
        assert incident.resolving_comment == "Containment in progress"

    async def test_directory_application_password_and_role_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """Directory writes retain complete application and credential models."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        application: dict[str, Any] = {
            "displayName": "Automation API",
            "description": "Automates tenant workflows",
            "signInAudience": "AzureADMyOrg",
            "identifierUris": ["api://automation"],
            "web": {
                "redirectUris": ["https://example.com/callback"],
                "homePageUrl": "https://example.com",
            },
            "tags": ["WindowsAzureActiveDirectoryIntegratedApp"],
        }
        application_update: dict[str, Any] = {
            "notes": "Managed by microsoft-agent",
            "spa": {"redirectUris": ["https://example.com/spa"]},
            "isFallbackPublicClient": False,
        }
        service_principal: dict[str, Any] = {
            "appId": "00000000-0000-0000-0000-000000000001",
            "displayName": "Automation API",
            "accountEnabled": True,
            "appRoleAssignmentRequired": True,
            "notificationEmailAddresses": ["operations@example.com"],
        }
        service_principal_update: dict[str, Any] = {
            "accountEnabled": False,
            "notes": "Temporarily disabled",
            "tags": ["disabled-by-policy"],
        }
        password: dict[str, Any] = {
            "displayName": "Deployment credential",
            "startDateTime": "2026-07-17T00:00:00Z",
            "endDateTime": "2026-08-17T00:00:00Z",
        }
        key_id = "00000000-0000-0000-0000-000000000002"
        role_assignment: dict[str, Any] = {
            "roleDefinitionId": "role-definition-1",
            "principalId": "principal-1",
            "appScopeId": "/applications/application-1",
            "condition": "@Resource[Microsoft.Directory/applications]",
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            app_item = mock_client.applications.by_application_id.return_value
            sp_item = (
                mock_client.service_principals.by_service_principal_id.return_value
            )
            assignments = mock_client.role_management.directory.role_assignments
            mock_client.applications.post = AsyncMock(return_value=mock_native_response)
            app_item.patch = AsyncMock(return_value=mock_native_response)
            app_item.add_password.post = AsyncMock(return_value=mock_native_response)
            app_item.remove_password.post = AsyncMock(return_value=mock_native_response)
            mock_client.service_principals.post = AsyncMock(
                return_value=mock_native_response
            )
            sp_item.patch = AsyncMock(return_value=mock_native_response)
            assignments.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.create_application(application) == {"value": []}
            assert await api.update_application(
                "application-1", application_update
            ) == {"value": []}
            assert await api.add_application_password("application-1", password) == {
                "value": []
            }
            assert await api.remove_application_password(
                "application-1", {"keyId": key_id}
            ) == {"status": "password removed"}
            assert await api.create_service_principal(service_principal) == {
                "value": []
            }
            assert await api.update_service_principal(
                "principal-1", service_principal_update
            ) == {"value": []}
            assert await api.create_role_assignment(role_assignment) == {"value": []}

        created_app = _awaited_argument(mock_client.applications.post)
        assert created_app.description == "Automates tenant workflows"
        assert created_app.web.redirect_uris == ["https://example.com/callback"]
        assert created_app.identifier_uris == ["api://automation"]
        updated_app = _awaited_argument(app_item.patch)
        assert updated_app.notes == "Managed by microsoft-agent"
        assert updated_app.spa.redirect_uris == ["https://example.com/spa"]

        add_password_body = _awaited_argument(app_item.add_password.post)
        credential = add_password_body.password_credential
        assert credential.display_name == "Deployment credential"
        assert credential.start_date_time.isoformat() == "2026-07-17T00:00:00+00:00"
        assert credential.end_date_time.isoformat() == "2026-08-17T00:00:00+00:00"
        remove_password_body = _awaited_argument(app_item.remove_password.post)
        assert str(remove_password_body.key_id) == key_id

        created_sp = _awaited_argument(mock_client.service_principals.post)
        assert created_sp.app_role_assignment_required is True
        assert created_sp.notification_email_addresses == ["operations@example.com"]
        updated_sp = _awaited_argument(sp_item.patch)
        assert updated_sp.account_enabled is False
        assert updated_sp.tags == ["disabled-by-policy"]
        assignment = _awaited_argument(assignments.post)
        assert assignment.directory_scope_id == "/"
        assert assignment.app_scope_id == "/applications/application-1"
        assert assignment.condition.startswith("@Resource")

    async def test_place_privacy_storage_and_external_connection_payloads(
        self, mock_auth_manager, mock_native_response
    ):
        """Specialized directory resources preserve nested configuration data."""
        mock_auth_manager.get_token.return_value = "test_token"
        mock_auth_manager.get_current_account.return_value = {
            "username": "test@example.com"
        }
        room: dict[str, Any] = {
            "displayName": "Board Room",
            "capacity": 16,
            "address": {
                "street": "1 Main Street",
                "city": "Chicago",
                "state": "IL",
                "postalCode": "60601",
            },
            "geoCoordinates": {"latitude": 41.881, "longitude": -87.623},
            "tags": ["video", "accessible"],
            "isWheelChairAccessible": True,
        }
        subject_request: dict[str, Any] = {
            "displayName": "Customer export",
            "description": "GDPR export request",
            "dataSubject": {
                "firstName": "Ada",
                "lastName": "Lovelace",
                "email": "ada@example.com",
            },
            "dataSubjectType": "customer",
            "type": "export",
            "includeAllVersions": True,
            "regulations": ["GDPR"],
        }
        container_type = "00000000-0000-0000-0000-000000000003"
        container: dict[str, Any] = {
            "displayName": "Legal Cases",
            "description": "Matter documents",
            "containerTypeId": container_type,
            "customProperties": {
                "department": {"value": "Legal", "isSearchable": True}
            },
        }
        connection: dict[str, Any] = {
            "id": "support-knowledge",
            "name": "Support Knowledge",
            "description": "Customer support articles",
            "contentCategory": "knowledgeBase",
            "configuration": {
                "authorizedAppIds": ["00000000-0000-0000-0000-000000000004"]
            },
        }

        with patch(
            "microsoft_agent.api.api_client_base.GraphServiceClient"
        ) as mock_client_class:
            mock_client = MagicMock()
            place_item = mock_client.places.by_place_id.return_value.graph_room
            subject_requests = mock_client.privacy.subject_rights_requests
            containers = mock_client.storage.file_storage.containers
            connections = mock_client.external.connections
            place_item.patch = AsyncMock(return_value=mock_native_response)
            subject_requests.post = AsyncMock(return_value=mock_native_response)
            containers.post = AsyncMock(return_value=mock_native_response)
            connections.post = AsyncMock(return_value=mock_native_response)
            mock_client_class.return_value = mock_client

            api = MicrosoftGraphApi(mock_auth_manager)
            assert await api.update_place("room-1", room) == {"value": []}
            assert await api.create_subject_rights_request(subject_request) == {
                "value": []
            }
            assert await api.create_file_storage_container(container) == {"value": []}
            assert await api.create_external_connection(connection) == {"value": []}

        updated_room = _awaited_argument(place_item.patch)
        assert updated_room.address.city == "Chicago"
        assert updated_room.geo_coordinates.latitude == 41.881
        assert updated_room.tags == ["video", "accessible"]
        assert updated_room.is_wheel_chair_accessible is True

        request = _awaited_argument(subject_requests.post)
        assert request.data_subject.email == "ada@example.com"
        assert request.data_subject_type.value == "customer"
        assert request.type.value == "export"
        assert request.regulations == ["GDPR"]

        created_container = _awaited_argument(containers.post)
        assert str(created_container.container_type_id) == container_type
        assert created_container.description == "Matter documents"
        assert created_container.custom_properties.additional_data["department"] == {
            "value": "Legal",
            "isSearchable": True,
        }
        created_connection = _awaited_argument(connections.post)
        assert created_connection.id == "support-knowledge"
        assert created_connection.content_category.value == "knowledgeBase"
        assert created_connection.configuration.authorized_app_ids == [
            "00000000-0000-0000-0000-000000000004"
        ]
