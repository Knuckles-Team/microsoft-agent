"""MCP tools for reports operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_reports_tools(mcp: FastMCP):
    @mcp.tool(tags={"reports"})
    async def microsoft_reports(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_email_activity_report', 'get_mailbox_usage_report', 'get_office365_active_users', 'get_sharepoint_activity_report', 'get_teams_user_activity', 'get_onedrive_usage_report'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft reports operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_email_activity_report":
            return client.get_email_activity_report(**kwargs)
        if action == "get_mailbox_usage_report":
            return client.get_mailbox_usage_report(**kwargs)
        if action == "get_office365_active_users":
            return client.get_office365_active_users(**kwargs)
        if action == "get_sharepoint_activity_report":
            return client.get_sharepoint_activity_report(**kwargs)
        if action == "get_teams_user_activity":
            return client.get_teams_user_activity(**kwargs)
        if action == "get_onedrive_usage_report":
            return client.get_onedrive_usage_report(**kwargs)
        raise ValueError(f"Unknown action: {action}")
