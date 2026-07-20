"""MCP tools for admin operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_ADMIN_ACTIONS = (
    "list_service_health",
    "get_service_health",
    "list_service_health_issues",
    "get_service_health_issue",
    "list_service_update_messages",
    "get_service_update_message",
    "get_admin_sharepoint",
    "update_admin_sharepoint",
    "list_delegated_admin_relationships",
    "get_delegated_admin_relationship",
)


def register_admin_tools(mcp: FastMCP):
    @mcp.tool(tags={"admin"})
    async def microsoft_admin(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_service_health', 'get_service_health', 'list_service_health_issues', 'get_service_health_issue', 'list_service_update_messages', 'get_service_update_message', 'get_admin_sharepoint', 'update_admin_sharepoint', 'list_delegated_admin_relationships', 'get_delegated_admin_relationship'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
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
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _ADMIN_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_service_health":
            return await invoke_client_method(client.list_service_health, **kwargs)
        if action == "get_service_health":
            return await invoke_client_method(client.get_service_health, **kwargs)
        if action == "list_service_health_issues":
            return await invoke_client_method(
                client.list_service_health_issues, **kwargs
            )
        if action == "get_service_health_issue":
            return await invoke_client_method(client.get_service_health_issue, **kwargs)
        if action == "list_service_update_messages":
            return await invoke_client_method(
                client.list_service_update_messages, **kwargs
            )
        if action == "get_service_update_message":
            return await invoke_client_method(
                client.get_service_update_message, **kwargs
            )
        if action == "get_admin_sharepoint":
            return await invoke_client_method(client.get_admin_sharepoint, **kwargs)
        if action == "update_admin_sharepoint":
            return await invoke_client_method(client.update_admin_sharepoint, **kwargs)
        if action == "list_delegated_admin_relationships":
            return await invoke_client_method(
                client.list_delegated_admin_relationships, **kwargs
            )
        if action == "get_delegated_admin_relationship":
            return await invoke_client_method(
                client.get_delegated_admin_relationship, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
