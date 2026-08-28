"""MCP tools for search operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_SEARCH_ACTIONS = ("search_query", "search_tools")


def register_search_tools(mcp: FastMCP):
    @mcp.tool(tags={"search"})
    async def microsoft_search(
        action: str = Field(
            description="Action to perform. Must be one of: 'search_query', 'search_tools'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft search operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _SEARCH_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "search_query":
            return await invoke_client_method(client.search_query, **kwargs)
        if action == "search_tools":
            return await invoke_client_method(client.search_tools, **kwargs)
        raise ValueError(f"Unknown action: {action}")
