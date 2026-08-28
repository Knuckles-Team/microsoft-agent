"""MCP tools for directory operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_DIRECTORY_ACTIONS = (
    "list_directory_objects",
    "get_directory_object",
    "list_directory_roles",
    "get_directory_role",
    "list_directory_role_templates",
    "list_deleted_items",
    "restore_deleted_item",
    "list_role_definitions",
    "get_role_definition",
    "list_role_assignments",
    "get_role_assignment",
    "create_role_assignment",
)


def register_directory_tools(mcp: FastMCP):
    @mcp.tool(tags={"directory"})
    async def microsoft_directory(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_directory_objects', 'get_directory_object', 'list_directory_roles', 'get_directory_role', 'list_directory_role_templates', 'list_deleted_items', 'restore_deleted_item', 'list_role_definitions', 'get_role_definition', 'list_role_assignments', 'get_role_assignment', 'create_role_assignment'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft directory operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _DIRECTORY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
