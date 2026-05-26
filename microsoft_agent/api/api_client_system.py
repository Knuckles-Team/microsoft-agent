import os
import sys
from typing import Any

from microsoft_agent.auth import AuthManager

CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite",
    "Tasks.ReadWrite",
    "Contacts.ReadWrite",
    "Group.ReadWrite.All",
    "Directory.Read.All",
    "Sites.Read.All",
    "Chat.Read",
    "ChatMessage.Read.All",
    "ChannelMessage.Read.All",
    "ServiceHealth.Read.All",
    "ServiceMessage.Read.All",
    "Domain.ReadWrite.All",
    "Organization.ReadWrite.All",
    "OnlineMeetings.ReadWrite",
    "CallRecords.Read.All",
    "Presence.Read.All",
    "User.Invite.All",
    "SecurityEvents.ReadWrite.All",
    "SecurityIncident.ReadWrite.All",
    "ThreatHunting.Read.All",
    "AuditLog.Read.All",
    "Reports.Read.All",
    "Application.ReadWrite.All",
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "IdentityRiskEvent.Read.All",
    "IdentityRiskyUser.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory",
    "EntitlementManagement.Read.All",
    "AccessReview.Read.All",
    "LifecycleWorkflows.Read.All",
]

# Only create global auth_manager if not in test mode
auth_manager: AuthManager | None
if not os.environ.get("TESTING"):
    auth_manager = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)
else:
    auth_manager = None

from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiSystem(MicrosoftGraphApiBase):
    def login(self, force: bool = False) -> str:
        """Authenticate with Microsoft."""
        if not force:
            token = self.auth_manager.get_token()
            if token:
                return "Already authenticated."

        def callback(msg):
            print(msg, file=sys.stderr)

        return self.auth_manager.acquire_token_by_device_code(callback=callback)

    def logout(self) -> str:
        """Logout."""
        self.auth_manager.logout()
        return "Logged out."

    def verify_login(self) -> str:
        """Verify login status."""
        token = self.auth_manager.get_token()
        if token:
            account = self.auth_manager.get_current_account()
            username = account.get("username") if account else "Unknown"
            return f"Authenticated as {username}"
        return "Not authenticated."

    def list_accounts(self) -> list[dict[str, Any]]:
        """List accounts."""
        return self.auth_manager.list_accounts()

    def search_tools(self, query: str, limit: int = 10) -> list[str]:
        """Search methods in this class."""

        matches = []
        for name in dir(self):
            if name.startswith("_"):
                continue
            if query.lower() in name.lower():
                matches.append(name)
            if len(matches) >= limit:
                break
        return matches
