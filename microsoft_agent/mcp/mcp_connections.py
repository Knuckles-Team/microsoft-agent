"""MCP tools for connections operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_connections_tools(mcp: FastMCP):
    @mcp.tool(tags={"connections"})
    async def microsoft_connections(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_external_connections', 'get_external_connection', 'create_external_connection', 'delete_external_connection'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft connections operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_external_connections":
            return await run_blocking(client.list_external_connections, **kwargs)
        if action == "get_external_connection":
            return await run_blocking(client.get_external_connection, **kwargs)
        if action == "create_external_connection":
            return await run_blocking(client.create_external_connection, **kwargs)
        if action == "delete_external_connection":
            return await run_blocking(client.delete_external_connection, **kwargs)
        raise ValueError(f"Unknown action: {action}")
