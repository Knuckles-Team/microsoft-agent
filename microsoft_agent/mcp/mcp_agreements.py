"""MCP tools for agreements operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_agreements_tools(mcp: FastMCP):
    @mcp.tool(tags={"agreements"})
    async def microsoft_agreements(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_agreements', 'get_agreement', 'create_agreement', 'delete_agreement'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft agreements operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_agreements":
            return await run_blocking(client.list_agreements, **kwargs)
        if action == "get_agreement":
            return await run_blocking(client.get_agreement, **kwargs)
        if action == "create_agreement":
            return await run_blocking(client.create_agreement, **kwargs)
        if action == "delete_agreement":
            return await run_blocking(client.delete_agreement, **kwargs)
        raise ValueError(f"Unknown action: {action}")
