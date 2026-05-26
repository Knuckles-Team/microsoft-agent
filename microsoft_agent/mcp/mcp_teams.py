"""MCP tools for teams operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_teams_tools(mcp: FastMCP):
    @mcp.tool(tags={"teams"})
    async def microsoft_teams(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_team', 'get_team_channel'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft teams operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_team":
            return client.get_team(**kwargs)
        if action == "get_team_channel":
            return client.get_team_channel(**kwargs)
        raise ValueError(f"Unknown action: {action}")
