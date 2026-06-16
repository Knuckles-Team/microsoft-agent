"""MCP tools for groups operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_groups_tools(mcp: FastMCP):
    @mcp.tool(tags={"groups"})
    async def microsoft_groups(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_groups', 'get_group', 'create_group', 'update_group', 'delete_group', 'list_group_members', 'add_group_member', 'remove_group_member', 'list_group_owners', 'list_group_conversations', 'list_group_drives'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft groups operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_groups":
            return await run_blocking(client.list_groups, **kwargs)
        if action == "get_group":
            return await run_blocking(client.get_group, **kwargs)
        if action == "create_group":
            return await run_blocking(client.create_group, **kwargs)
        if action == "update_group":
            return await run_blocking(client.update_group, **kwargs)
        if action == "delete_group":
            return await run_blocking(client.delete_group, **kwargs)
        if action == "list_group_members":
            return await run_blocking(client.list_group_members, **kwargs)
        if action == "add_group_member":
            return await run_blocking(client.add_group_member, **kwargs)
        if action == "remove_group_member":
            return await run_blocking(client.remove_group_member, **kwargs)
        if action == "list_group_owners":
            return await run_blocking(client.list_group_owners, **kwargs)
        if action == "list_group_conversations":
            return await run_blocking(client.list_group_conversations, **kwargs)
        if action == "list_group_drives":
            return await run_blocking(client.list_group_drives, **kwargs)
        raise ValueError(f"Unknown action: {action}")
