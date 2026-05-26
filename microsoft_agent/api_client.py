"""Microsoft API Client.

CONCEPT:ECO-4.1
"""

import os

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

from microsoft_agent.api.api_client_admin import MicrosoftGraphApiAdmin
from microsoft_agent.api.api_client_apps import MicrosoftGraphApiApps
from microsoft_agent.api.api_client_calendar import MicrosoftGraphApiCalendar
from microsoft_agent.api.api_client_directory import MicrosoftGraphApiDirectory
from microsoft_agent.api.api_client_drive import MicrosoftGraphApiDrive
from microsoft_agent.api.api_client_mail import MicrosoftGraphApiMail
from microsoft_agent.api.api_client_other import MicrosoftGraphApiOther
from microsoft_agent.api.api_client_system import MicrosoftGraphApiSystem


class MicrosoftGraphApi(
    MicrosoftGraphApiSystem,
    MicrosoftGraphApiMail,
    MicrosoftGraphApiCalendar,
    MicrosoftGraphApiDrive,
    MicrosoftGraphApiDirectory,
    MicrosoftGraphApiApps,
    MicrosoftGraphApiAdmin,
    MicrosoftGraphApiOther,
):
    pass
