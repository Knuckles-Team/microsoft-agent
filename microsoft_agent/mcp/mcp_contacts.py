"""MCP tools for contacts operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_contacts_tools(mcp: FastMCP):
    @mcp.tool(tags={"contacts"})
    async def microsoft_contacts(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_outlook_contact', 'create_outlook_contact', 'update_outlook_contact', 'delete_outlook_contact'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft contacts operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_outlook_contact":
            return client.get_outlook_contact(**kwargs)
        if action == "create_outlook_contact":
            return client.create_outlook_contact(**kwargs)
        if action == "update_outlook_contact":
            return client.update_outlook_contact(**kwargs)
        if action == "delete_outlook_contact":
            return client.delete_outlook_contact(**kwargs)
        raise ValueError(f"Unknown action: {action}")
