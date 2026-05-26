"""MCP tools for directory operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_directory_tools(mcp: FastMCP):
    @mcp.tool(tags={"directory"})
    async def microsoft_directory(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_directory_objects', 'get_directory_object', 'list_directory_roles', 'get_directory_role', 'list_directory_role_templates', 'list_deleted_items', 'restore_deleted_item', 'list_role_definitions', 'get_role_definition', 'list_role_assignments', 'get_role_assignment', 'create_role_assignment'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
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
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_directory_objects":
            return client.list_directory_objects(**kwargs)
        if action == "get_directory_object":
            return client.get_directory_object(**kwargs)
        if action == "list_directory_roles":
            return client.list_directory_roles(**kwargs)
        if action == "get_directory_role":
            return client.get_directory_role(**kwargs)
        if action == "list_directory_role_templates":
            return client.list_directory_role_templates(**kwargs)
        if action == "list_deleted_items":
            return client.list_deleted_items(**kwargs)
        if action == "restore_deleted_item":
            return client.restore_deleted_item(**kwargs)
        if action == "list_role_definitions":
            return client.list_role_definitions(**kwargs)
        if action == "get_role_definition":
            return client.get_role_definition(**kwargs)
        if action == "list_role_assignments":
            return client.list_role_assignments(**kwargs)
        if action == "get_role_assignment":
            return client.get_role_assignment(**kwargs)
        if action == "create_role_assignment":
            return client.create_role_assignment(**kwargs)
        raise ValueError(f"Unknown action: {action}")
