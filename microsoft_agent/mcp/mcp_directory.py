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
            ctx.info("Executing tool...")
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

        if action == "list_directory_objects":
            return await invoke_client_method(client.list_directory_objects, **kwargs)
        if action == "get_directory_object":
            return await invoke_client_method(client.get_directory_object, **kwargs)
        if action == "list_directory_roles":
            return await invoke_client_method(client.list_directory_roles, **kwargs)
        if action == "get_directory_role":
            return await invoke_client_method(client.get_directory_role, **kwargs)
        if action == "list_directory_role_templates":
            return await invoke_client_method(
                client.list_directory_role_templates, **kwargs
            )
        if action == "list_deleted_items":
            return await invoke_client_method(client.list_deleted_items, **kwargs)
        if action == "restore_deleted_item":
            return await invoke_client_method(client.restore_deleted_item, **kwargs)
        if action == "list_role_definitions":
            return await invoke_client_method(client.list_role_definitions, **kwargs)
        if action == "get_role_definition":
            return await invoke_client_method(client.get_role_definition, **kwargs)
        if action == "list_role_assignments":
            return await invoke_client_method(client.list_role_assignments, **kwargs)
        if action == "get_role_assignment":
            return await invoke_client_method(client.get_role_assignment, **kwargs)
        if action == "create_role_assignment":
            return await invoke_client_method(client.create_role_assignment, **kwargs)
        raise ValueError(f"Unknown action: {action}")
