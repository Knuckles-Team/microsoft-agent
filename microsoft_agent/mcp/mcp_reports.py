"""MCP tools for reports operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_REPORTS_ACTIONS = (
    "get_email_activity_report",
    "get_mailbox_usage_report",
    "get_office365_active_users",
    "get_sharepoint_activity_report",
    "get_teams_user_activity",
    "get_onedrive_usage_report",
)


def register_reports_tools(mcp: FastMCP):
    @mcp.tool(tags={"reports"})
    async def microsoft_reports(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_email_activity_report', 'get_mailbox_usage_report', 'get_office365_active_users', 'get_sharepoint_activity_report', 'get_teams_user_activity', 'get_onedrive_usage_report'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft reports operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _REPORTS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_email_activity_report":
            return await invoke_client_method(
                client.get_email_activity_report, **kwargs
            )
        if action == "get_mailbox_usage_report":
            return await invoke_client_method(client.get_mailbox_usage_report, **kwargs)
        if action == "get_office365_active_users":
            return await invoke_client_method(
                client.get_office365_active_users, **kwargs
            )
        if action == "get_sharepoint_activity_report":
            return await invoke_client_method(
                client.get_sharepoint_activity_report, **kwargs
            )
        if action == "get_teams_user_activity":
            return await invoke_client_method(client.get_teams_user_activity, **kwargs)
        if action == "get_onedrive_usage_report":
            return await invoke_client_method(
                client.get_onedrive_usage_report, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
