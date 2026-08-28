"""MCP tools for applications operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_APPLICATIONS_ACTIONS = (
    "list_applications",
    "get_application",
    "create_application",
    "update_application",
    "delete_application",
    "add_application_password",
    "remove_application_password",
    "list_service_principals",
    "get_service_principal",
    "create_service_principal",
    "update_service_principal",
    "delete_service_principal",
)


def register_applications_tools(mcp: FastMCP):
    @mcp.tool(tags={"applications"})
    async def microsoft_applications(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_applications', 'get_application', 'create_application', 'update_application', 'delete_application', 'add_application_password', 'remove_application_password', 'list_service_principals', 'get_service_principal', 'create_service_principal', 'update_service_principal', 'delete_service_principal'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft applications operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _APPLICATIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
