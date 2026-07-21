"""Pytest configuration and fixtures for Microsoft Agent tests.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

import atexit
import os
from datetime import datetime
from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from msgraph import GraphServiceClient

# Globally disable atexit hook registration during tests to prevent real keyring
# calls during interpreter exit.
atexit.register = lambda fn, *args, **kwargs: fn

# Set TESTING environment variable to prevent global auth_manager creation
os.environ["TESTING"] = "1"


@pytest.fixture(autouse=True, scope="session")
def mock_system_keyring():
    """Globally mock keyring to prevent database/SecretService keyring hangs in
    headless CI/test environments."""
    with (
        patch("keyring.get_password", return_value=None),
        patch("keyring.set_password", return_value=None),
        patch("keyring.delete_password", return_value=None),
    ):
        yield


# ============================================================================
# GRAPH API RESPONSE FACTORIES
# ============================================================================


class GraphApiResponseFactory:
    """Factory for creating Microsoft Graph API response objects."""

    @staticmethod
    def create_user(user_id: str = "user123", **kwargs) -> dict[str, Any]:
        """Create a mock user object."""
        defaults = {
            "id": user_id,
            "displayName": "Test User",
            "mail": "test@example.com",
            "userPrincipalName": "test@example.com",
            "givenName": "Test",
            "surname": "User",
            "jobTitle": "Developer",
            "department": "Engineering",
            "officeLocation": "Building 1",
            "businessPhones": ["+1-555-0100"],
            "mobilePhone": "+1-555-0101",
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_message(message_id: str = "msg123", **kwargs) -> dict[str, Any]:
        """Create a mock message object."""
        defaults = {
            "id": message_id,
            "subject": "Test Subject",
            "body": {"content": "Test content", "contentType": "Text"},
            "from": {
                "emailAddress": {"address": "sender@example.com", "name": "Sender Name"}
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": "recipient@example.com",
                        "name": "Recipient Name",
                    }
                }
            ],
            "receivedDateTime": datetime.now().isoformat(),
            "sentDateTime": datetime.now().isoformat(),
            "isRead": False,
            "isDraft": False,
            "hasAttachments": False,
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_calendar_event(event_id: str = "event123", **kwargs) -> dict[str, Any]:
        """Create a mock calendar event object."""
        defaults = {
            "id": event_id,
            "subject": "Test Meeting",
            "start": {"dateTime": "2024-01-01T10:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2024-01-01T11:00:00", "timeZone": "UTC"},
            "location": {"displayName": "Conference Room"},
            "attendees": [
                {
                    "emailAddress": {
                        "address": "attendee@example.com",
                        "name": "Attendee Name",
                    },
                    "type": "required",
                }
            ],
            "isAllDay": False,
            "showAs": "busy",
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_group(group_id: str = "group123", **kwargs) -> dict[str, Any]:
        """Create a mock group object."""
        defaults = {
            "id": group_id,
            "displayName": "Test Group",
            "description": "Test group description",
            "mail": "testgroup@example.com",
            "mailEnabled": True,
            "securityEnabled": False,
            "groupTypes": ["Unified"],
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_drive_item(item_id: str = "item123", **kwargs) -> dict[str, Any]:
        """Create a mock drive item object."""
        defaults = {
            "id": item_id,
            "name": "testfile.txt",
            "size": 1024,
            "file": {"mimeType": "text/plain"},
            "folder": None,
            "createdDateTime": datetime.now().isoformat(),
            "lastModifiedDateTime": datetime.now().isoformat(),
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_attachment(attachment_id: str = "att123", **kwargs) -> dict[str, Any]:
        """Create a mock attachment object."""
        defaults = {
            "id": attachment_id,
            "name": "attachment.txt",
            "size": 512,
            "contentType": "text/plain",
            "isInline": False,
        }
        defaults.update(kwargs)
        return defaults

    @staticmethod
    def create_collection(
        items: list[dict[str, Any]], next_link: str | None = None
    ) -> dict[str, Any]:
        """Create a mock Graph API collection response."""
        response: dict[str, Any] = {"value": items}
        if next_link:
            response["@odata.nextLink"] = next_link
        return response


# ============================================================================
# PARAMETERIZED TEST HELPERS
# ============================================================================


class ParameterizedTestHelper:
    """Helper class for parameterized testing of similar operations."""

    @staticmethod
    def generate_test_cases(
        base_data: dict[str, Any], variations: list[dict[str, Any]]
    ) -> list[dict[str, Any]]:
        """Generate test cases by applying variations to base data."""
        test_cases = []
        for variation in variations:
            test_case = base_data.copy()
            test_case.update(variation)
            test_cases.append(test_case)
        return test_cases

    @staticmethod
    def create_error_scenarios(success_response: Any) -> dict[str, Any]:
        """Create common error scenarios for testing."""
        return {
            "success": success_response,
            "network_error": Exception("Network error"),
            "auth_error": Exception("Authentication failed"),
            "not_found": Exception("Resource not found"),
            "permission_denied": Exception("Permission denied"),
            "rate_limit": Exception("Rate limit exceeded"),
            "server_error": Exception("Internal server error"),
        }


# ============================================================================
# MOCK GRAPH CLIENT BUILDER
# ============================================================================


class MockGraphClientBuilder:
    """Builder for creating comprehensive mock Graph clients."""

    def __init__(self):
        self.client = MagicMock(spec=GraphServiceClient)
        self._setup_basic_structure()

    def _setup_basic_structure(self):
        """Setup basic Graph client structure."""
        self.client.me = MagicMock()
        self.client.users = MagicMock()
        self.client.groups = MagicMock()
        self.client.drives = MagicMock()
        self.client.calendars = MagicMock()

    def with_user_operations(self) -> "MockGraphClientBuilder":
        """Add mock user operations."""
        self.client.me.get = AsyncMock(
            return_value=MagicMock(json=lambda: GraphApiResponseFactory.create_user())
        )
        self.client.users.by_id = MagicMock(
            return_value=MagicMock(
                get=AsyncMock(
                    return_value=MagicMock(
                        json=lambda: GraphApiResponseFactory.create_user()
                    )
                )
            )
        )
        return self

    def with_mail_operations(self) -> "MockGraphClientBuilder":
        """Add mock mail operations."""
        self.client.me.messages = MagicMock()
        self.client.me.messages.get = AsyncMock(
            return_value=MagicMock(
                json=lambda: GraphApiResponseFactory.create_collection(
                    [GraphApiResponseFactory.create_message()]
                )
            )
        )
        self.client.me.messages.by_id = MagicMock(
            return_value=MagicMock(
                get=AsyncMock(
                    return_value=MagicMock(
                        json=lambda: GraphApiResponseFactory.create_message()
                    )
                ),
                delete=AsyncMock(),
                patch=AsyncMock(
                    return_value=MagicMock(
                        json=lambda: GraphApiResponseFactory.create_message()
                    )
                ),
                move=AsyncMock(
                    return_value=MagicMock(
                        json=lambda: GraphApiResponseFactory.create_message()
                    )
                ),
            )
        )
        self.client.me.send_mail = AsyncMock()
        return self

    def with_calendar_operations(self) -> "MockGraphClientBuilder":
        """Add mock calendar operations."""
        self.client.me.calendar = MagicMock()
        self.client.me.calendar.events = MagicMock()
        self.client.me.calendar.events.get = AsyncMock(
            return_value=MagicMock(
                json=lambda: GraphApiResponseFactory.create_collection(
                    [GraphApiResponseFactory.create_calendar_event()]
                )
            )
        )
        self.client.me.calendar.events.post = AsyncMock(
            return_value=MagicMock(
                json=lambda: GraphApiResponseFactory.create_calendar_event()
            )
        )
        return self

    def with_group_operations(self) -> "MockGraphClientBuilder":
        """Add mock group operations."""
        self.client.groups.get = AsyncMock(
            return_value=MagicMock(
                json=lambda: GraphApiResponseFactory.create_collection(
                    [GraphApiResponseFactory.create_group()]
                )
            )
        )
        self.client.groups.post = AsyncMock(
            return_value=MagicMock(json=lambda: GraphApiResponseFactory.create_group())
        )
        return self

    def with_drive_operations(self) -> "MockGraphClientBuilder":
        """Add mock drive operations."""
        self.client.me.drive = MagicMock()
        self.client.me.drive.root = MagicMock()
        self.client.me.drive.root.children = MagicMock()
        self.client.me.drive.root.children.get = AsyncMock(
            return_value=MagicMock(
                json=lambda: GraphApiResponseFactory.create_collection(
                    [GraphApiResponseFactory.create_drive_item()]
                )
            )
        )
        return self

    def with_error_handling(
        self, error: Exception | None = None
    ) -> "MockGraphClientBuilder":
        """Add error handling to operations."""
        if error:
            self.client.me.get = AsyncMock(side_effect=error)
            self.client.me.messages.get = AsyncMock(side_effect=error)
        return self

    def build(self) -> MagicMock:
        """Build the mock client."""
        return self.client


# ============================================================================
# ENHANCED FIXTURES
# ============================================================================


@pytest.fixture
def graph_response_factory():
    """Graph API response factory fixture."""
    return GraphApiResponseFactory()


@pytest.fixture
def parameterized_test_helper():
    """Parameterized test helper fixture."""
    return ParameterizedTestHelper()


@pytest.fixture
def mock_graph_client_builder():
    """Mock Graph client builder fixture."""
    return MockGraphClientBuilder()


@pytest.fixture
def mock_comprehensive_graph_client():
    """Comprehensive mock Graph client with all operations."""
    return (
        MockGraphClientBuilder()
        .with_user_operations()
        .with_mail_operations()
        .with_calendar_operations()
        .with_group_operations()
        .with_drive_operations()
        .build()
    )


@pytest.fixture
def sample_graph_user_data():
    """Sample Graph user data for testing."""
    return GraphApiResponseFactory.create_user(
        user_id="user123",
        displayName="John Doe",
        mail="john.doe@example.com",
        jobTitle="Software Engineer",
    )


@pytest.fixture
def sample_graph_message_data():
    """Sample Graph message data for testing."""
    return GraphApiResponseFactory.create_message(
        message_id="msg123", subject="Test Subject", isRead=False
    )


@pytest.fixture
def sample_graph_calendar_data():
    """Sample Graph calendar event data for testing."""
    return GraphApiResponseFactory.create_calendar_event(
        event_id="event123", subject="Team Meeting"
    )


@pytest.fixture
def sample_graph_group_data():
    """Sample Graph group data for testing."""
    return GraphApiResponseFactory.create_group(
        group_id="group123", displayName="Engineering Team"
    )


@pytest.fixture
def sample_graph_drive_data():
    """Sample Graph drive item data for testing."""
    return GraphApiResponseFactory.create_drive_item(
        item_id="item123", name="document.txt"
    )


@pytest.fixture
def sample_graph_attachment_data():
    """Sample Graph attachment data for testing."""
    return GraphApiResponseFactory.create_attachment(
        attachment_id="att123", name="report.pdf"
    )


@pytest.fixture
def mock_graph_collection():
    """Mock Graph API collection response."""
    return GraphApiResponseFactory.create_collection(
        [GraphApiResponseFactory.create_user(user_id=f"user{i}") for i in range(5)]
    )


@pytest.fixture
def mock_graph_error_responses():
    """Mock Graph API error responses."""
    return ParameterizedTestHelper.create_error_scenarios(
        GraphApiResponseFactory.create_user()
    )


# ============================================================================
# SHARED AUTHENTICATION FIXTURES
# ============================================================================


@pytest.fixture
def mock_token_cache():
    """Mock token cache for testing."""
    return {}


@pytest.fixture
def mock_auth_manager():
    """Mock AuthManager for testing."""
    mock_auth = MagicMock()
    mock_auth.client_id = "test_client_id"
    mock_auth.authority = "https://login.microsoftonline.com/common"
    mock_auth.scopes = ["User.Read", "Mail.ReadWrite"]
    mock_auth.graph_base_url = "https://graph.microsoft.com/v1.0"
    mock_auth.graph_tls_profile = None
    mock_auth.graph_tls_profile_ref = None
    mock_auth.access_token = None
    mock_auth.selected_account_id = None
    mock_auth.get_token.return_value = None
    mock_auth.get_current_account.return_value = None
    mock_auth.get_token_details.return_value = None
    mock_auth.list_accounts.return_value = []
    mock_auth.login.return_value = "Authentication successful"
    mock_auth.logout.return_value = "Logged out."
    mock_auth.verify_login.return_value = "Not authenticated."
    return mock_auth


@pytest.fixture
def real_auth_manager():
    """Real AuthManager for testing auth module."""
    with patch("microsoft_agent.auth.msal.PublicClientApplication") as mock_msal:
        mock_app = MagicMock()
        mock_app.get_accounts.return_value = []
        mock_msal.return_value = mock_app

        with patch("microsoft_agent.auth.keyring.get_password", return_value=None):
            with patch("microsoft_agent.auth.keyring.set_password"):
                from microsoft_agent.auth import AuthManager

                auth = AuthManager(
                    client_id="test_client_id",
                    authority="https://login.microsoftonline.com/common",
                    scopes=["User.Read", "Mail.ReadWrite"],
                )
                # Reset state for clean test isolation
                auth.selected_account_id = None
                auth.access_token = None
                return auth


@pytest.fixture
def mock_graph_client():
    """Mock Microsoft Graph client for testing."""
    mock_client = MagicMock(spec=GraphServiceClient)
    mock_client.me = MagicMock()
    mock_client.users = MagicMock()
    return mock_client


@pytest.fixture
def mock_native_response():
    """Mock native response handler."""
    mock_response = MagicMock()
    mock_response.json.return_value = {"value": []}
    mock_response.raise_for_status = MagicMock()
    return mock_response


@pytest.fixture
def temp_token_cache_dir(tmp_path):
    """Temporary directory for token cache testing."""
    cache_dir = tmp_path / ".microsoft-agent"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


@pytest.fixture
def mock_environment():
    """Mock environment variables for testing."""
    original_env = os.environ.copy()
    os.environ["OIDC_CLIENT_ID"] = "test_client_id"
    yield
    os.environ.clear()
    os.environ.update(original_env)


@pytest.fixture
def sample_mail_data():
    """Sample mail data for testing."""
    return {
        "message": {
            "subject": "Test Subject",
            "body": {"content": "Test content", "contentType": "Text"},
            "toRecipients": [
                {"emailAddress": {"address": "test@example.com", "name": "Test User"}}
            ],
        },
        "saveToSentItems": True,
    }


@pytest.fixture
def sample_user_data():
    """Sample user data for testing."""
    return {"id": "user123", "displayName": "Test User", "mail": "test@example.com"}


@pytest.fixture
def sample_calendar_event_data():
    """Sample calendar event data for testing."""
    return {
        "subject": "Test Meeting",
        "start": {"dateTime": "2024-01-01T10:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2024-01-01T11:00:00", "timeZone": "UTC"},
    }


@pytest.fixture
def mock_query_params():
    """Mock query parameters for testing."""
    return {
        "$select": "id,subject",
        "$top": "10",
        "$filter": "isRead eq false",
        "$search": "test",
        "$orderby": "receivedDateTime desc",
    }


@pytest.fixture
def mock_device_flow():
    """Mock device code flow for testing."""
    return {
        "user_code": "TEST123",
        "device_code": "device_code_123",
        "verification_url": "https://microsoft.com/devicelogin",
        "message": "Please visit https://microsoft.com/devicelogin and enter code TEST123",
    }


@pytest.fixture
def mock_token_response():
    """Mock token response for testing."""
    return {
        "access_token": "test_access_token",
        "expires_in": 3600,
        "expires_on": 1234567890,
        "home_account_id": "account123",
    }


@pytest.fixture
def mock_account():
    """Mock account for testing."""
    return {"home_account_id": "account123", "username": "test@example.com"}


def create_mock_request_builder():
    """Create a mock request builder for testing."""
    mock_builder = MagicMock()
    mock_builder.get = AsyncMock(return_value=MagicMock(json=lambda: {"value": []}))
    mock_builder.post = AsyncMock(return_value=MagicMock(json=lambda: {"id": "123"}))
    mock_builder.patch = AsyncMock(return_value=MagicMock(json=lambda: {"id": "123"}))
    mock_builder.delete = AsyncMock()
    return mock_builder


@pytest.fixture
def mock_request_builder():
    """Mock request builder fixture."""
    return create_mock_request_builder()
