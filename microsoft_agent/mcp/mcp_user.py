"""MCP tools for user operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_user_tools(mcp: FastMCP):
    @mcp.tool(tags={"user"})
    async def microsoft_user(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_current_user', 'get_me'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft user operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_current_user":
            return client.get_current_user(**kwargs)
        if action == "get_me":
            return client.get_me(**kwargs)
        raise ValueError(f"Unknown action: {action}")
