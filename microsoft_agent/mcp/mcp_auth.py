"""MCP tools for auth operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_auth_tools(mcp: FastMCP):
    @mcp.tool(tags={"auth"})
    async def microsoft_auth(
        action: str = Field(
            description="Action to perform. Must be one of: 'login', 'logout', 'verify_login', 'list_accounts'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft auth operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "login":
            return await run_blocking(client.login, **kwargs)
        if action == "logout":
            return await run_blocking(client.logout, **kwargs)
        if action == "verify_login":
            return await run_blocking(client.verify_login, **kwargs)
        if action == "list_accounts":
            return await run_blocking(client.list_accounts, **kwargs)
        raise ValueError(f"Unknown action: {action}")
