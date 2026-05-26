"""MCP tools for print operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_print_tools(mcp: FastMCP):
    @mcp.tool(tags={"print"})
    async def microsoft_print(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_printers', 'get_printer', 'list_print_jobs', 'create_print_job', 'list_print_shares'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft print operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_printers":
            return client.list_printers(**kwargs)
        if action == "get_printer":
            return client.get_printer(**kwargs)
        if action == "list_print_jobs":
            return client.list_print_jobs(**kwargs)
        if action == "create_print_job":
            return client.create_print_job(**kwargs)
        if action == "list_print_shares":
            return client.list_print_shares(**kwargs)
        raise ValueError(f"Unknown action: {action}")
