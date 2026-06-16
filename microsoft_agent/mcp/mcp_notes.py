"""MCP tools for notes operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_notes_tools(mcp: FastMCP):
    @mcp.tool(tags={"notes"})
    async def microsoft_notes(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_onenote_page_content', 'create_onenote_page'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft notes operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_onenote_page_content":
            return await run_blocking(client.get_onenote_page_content, **kwargs)
        if action == "create_onenote_page":
            return await run_blocking(client.create_onenote_page, **kwargs)
        raise ValueError(f"Unknown action: {action}")
