"""Microsoft Agent.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

from scion_sdk.connector import Connector
from scion_sdk.connector_hooks import RegistrationContext

from microsoft_agent._version import __version__


def _register_kg_tools(context: RegistrationContext) -> None:
    """Register the knowledge-graph ingestion tool through the SDK's shared hook."""
    from microsoft_agent.mcp_server import register_kg_tools

    register_kg_tools(context.server)


connector = Connector(
    instructions="microsoft-agent MCP Server — Condensed Action-Routed Tools.",
    extra_registrations=(_register_kg_tools,),
)

__all__ = ["__version__", "connector"]
