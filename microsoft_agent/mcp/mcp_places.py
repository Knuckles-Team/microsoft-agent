"""MCP tools for places operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_places_tools(mcp: FastMCP):
    @mcp.tool(tags={"places"})
    async def microsoft_places(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_rooms', 'list_room_lists', 'get_place', 'update_place'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft places operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_rooms":
            return await run_blocking(client.list_rooms, **kwargs)
        if action == "list_room_lists":
            return await run_blocking(client.list_room_lists, **kwargs)
        if action == "get_place":
            return await run_blocking(client.get_place, **kwargs)
        if action == "update_place":
            return await run_blocking(client.update_place, **kwargs)
        raise ValueError(f"Unknown action: {action}")
