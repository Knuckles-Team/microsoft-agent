"""MCP tools for solutions operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_solutions_tools(mcp: FastMCP):
    @mcp.tool(tags={"solutions"})
    async def microsoft_solutions(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_booking_businesses', 'get_booking_business', 'list_booking_appointments', 'create_booking_appointment', 'list_virtual_events'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft solutions operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_booking_businesses":
            return client.list_booking_businesses(**kwargs)
        if action == "get_booking_business":
            return client.get_booking_business(**kwargs)
        if action == "list_booking_appointments":
            return client.list_booking_appointments(**kwargs)
        if action == "create_booking_appointment":
            return client.create_booking_appointment(**kwargs)
        if action == "list_virtual_events":
            return client.list_virtual_events(**kwargs)
        raise ValueError(f"Unknown action: {action}")
