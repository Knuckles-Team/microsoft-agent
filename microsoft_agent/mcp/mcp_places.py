"""MCP tools for places operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_PLACES_ACTIONS = ("list_rooms", "list_room_lists", "get_place", "update_place")


def register_places_tools(mcp: FastMCP):
    @mcp.tool(tags={"places"})
    async def microsoft_places(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_rooms', 'list_room_lists', 'get_place', 'update_place'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
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
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _PLACES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_rooms":
            return await invoke_client_method(client.list_rooms, **kwargs)
        if action == "list_room_lists":
            return await invoke_client_method(client.list_room_lists, **kwargs)
        if action == "get_place":
            return await invoke_client_method(client.get_place, **kwargs)
        if action == "update_place":
            return await invoke_client_method(client.update_place, **kwargs)
        raise ValueError(f"Unknown action: {action}")
