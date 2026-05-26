"""MCP tools for calendar operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_calendar_tools(mcp: FastMCP):
    @mcp.tool(tags={"calendar"})
    async def microsoft_calendar(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_calendar_events', 'get_calendar_event', 'create_calendar_event', 'update_calendar_event', 'delete_calendar_event', 'list_specific_calendar_events', 'get_specific_calendar_event', 'create_specific_calendar_event', 'update_specific_calendar_event', 'delete_specific_calendar_event', 'get_calendar_view', 'list_calendars', 'find_meeting_times'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft calendar operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_calendar_events":
            return client.list_calendar_events(**kwargs)
        if action == "get_calendar_event":
            return client.get_calendar_event(**kwargs)
        if action == "create_calendar_event":
            return client.create_calendar_event(**kwargs)
        if action == "update_calendar_event":
            return client.update_calendar_event(**kwargs)
        if action == "delete_calendar_event":
            return client.delete_calendar_event(**kwargs)
        if action == "list_specific_calendar_events":
            return client.list_specific_calendar_events(**kwargs)
        if action == "get_specific_calendar_event":
            return client.get_specific_calendar_event(**kwargs)
        if action == "create_specific_calendar_event":
            return client.create_specific_calendar_event(**kwargs)
        if action == "update_specific_calendar_event":
            return client.update_specific_calendar_event(**kwargs)
        if action == "delete_specific_calendar_event":
            return client.delete_specific_calendar_event(**kwargs)
        if action == "get_calendar_view":
            return client.get_calendar_view(**kwargs)
        if action == "list_calendars":
            return client.list_calendars(**kwargs)
        if action == "find_meeting_times":
            return client.find_meeting_times(**kwargs)
        raise ValueError(f"Unknown action: {action}")
