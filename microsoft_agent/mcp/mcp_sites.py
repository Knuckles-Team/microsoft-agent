"""MCP tools for sites operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_sites_tools(mcp: FastMCP):
    @mcp.tool(tags={"sites"})
    async def microsoft_sites(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_sites', 'get_site', 'get_sharepoint_site_by_path', 'get_sharepoint_sites_delta'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft sites operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_sites":
            return await run_blocking(client.list_sites, **kwargs)
        if action == "get_site":
            return await run_blocking(client.get_site, **kwargs)
        if action == "get_sharepoint_site_by_path":
            return await run_blocking(client.get_sharepoint_site_by_path, **kwargs)
        if action == "get_sharepoint_sites_delta":
            return await run_blocking(client.get_sharepoint_sites_delta, **kwargs)
        raise ValueError(f"Unknown action: {action}")
