"""MCP tools for admin operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_admin_tools(mcp: FastMCP):
    @mcp.tool(tags={"admin"})
    async def microsoft_admin(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_service_health', 'get_service_health', 'list_service_health_issues', 'get_service_health_issue', 'list_service_update_messages', 'get_service_update_message', 'get_admin_sharepoint', 'update_admin_sharepoint', 'list_delegated_admin_relationships', 'get_delegated_admin_relationship'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft admin operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_service_health":
            return client.list_service_health(**kwargs)
        if action == "get_service_health":
            return client.get_service_health(**kwargs)
        if action == "list_service_health_issues":
            return client.list_service_health_issues(**kwargs)
        if action == "get_service_health_issue":
            return client.get_service_health_issue(**kwargs)
        if action == "list_service_update_messages":
            return client.list_service_update_messages(**kwargs)
        if action == "get_service_update_message":
            return client.get_service_update_message(**kwargs)
        if action == "get_admin_sharepoint":
            return client.get_admin_sharepoint(**kwargs)
        if action == "update_admin_sharepoint":
            return client.update_admin_sharepoint(**kwargs)
        if action == "list_delegated_admin_relationships":
            return client.list_delegated_admin_relationships(**kwargs)
        if action == "get_delegated_admin_relationship":
            return client.get_delegated_admin_relationship(**kwargs)
        raise ValueError(f"Unknown action: {action}")
