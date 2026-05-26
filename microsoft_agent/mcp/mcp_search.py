"""MCP tools for search operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_search_tools(mcp: FastMCP):
    @mcp.tool(tags={"search"})
    async def microsoft_search(
        action: str = Field(
            description="Action to perform. Must be one of: 'search_query', 'search_tools'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft search operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "search_query":
            return client.search_query(**kwargs)
        if action == "search_tools":
            return client.search_tools(**kwargs)
        raise ValueError(f"Unknown action: {action}")
