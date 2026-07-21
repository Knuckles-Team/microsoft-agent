"""Microsoft API Client.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

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
