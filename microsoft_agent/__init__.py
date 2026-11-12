"""Microsoft Agent.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

from scion_sdk.connector import Connector
from scion_sdk.connector_hooks import RegistrationContext

from microsoft_agent._version import __version__
from microsoft_agent.api_client import MicrosoftGraphApi
from microsoft_agent.auth import get_client_dependency


def _register_extras(context: RegistrationContext) -> None:
    """Register the non-Graph tools and the side-effect policy through the SDK hook."""
    from microsoft_agent.mcp_server import register_extra_surface

    register_extra_surface(context.server)


connector = Connector(
    instructions="microsoft-agent MCP Server: one typed tool per Microsoft Graph operation.",
    api=MicrosoftGraphApi,
    get_client=get_client_dependency,
    extra_registrations=(_register_extras,),
)

__all__ = ["__version__", "connector"]
