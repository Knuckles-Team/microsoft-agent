"""MCP tools for notes operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_NOTES_ACTIONS = ("get_onenote_page_content", "create_onenote_page")


def register_notes_tools(mcp: FastMCP):
    @mcp.tool(tags={"notes"})
    async def microsoft_notes(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_onenote_page_content', 'create_onenote_page'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft notes operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _NOTES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_onenote_page_content":
            return await invoke_client_method(client.get_onenote_page_content, **kwargs)
        if action == "create_onenote_page":
            return await invoke_client_method(client.create_onenote_page, **kwargs)
        raise ValueError(f"Unknown action: {action}")
