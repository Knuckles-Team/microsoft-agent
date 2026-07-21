"""MCP tools for organization operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_ORGANIZATION_ACTIONS = (
    "list_organization",
    "get_organization",
    "update_organization",
    "get_org_branding",
    "update_org_branding",
)


def register_organization_tools(mcp: FastMCP):
    @mcp.tool(tags={"organization"})
    async def microsoft_organization(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_organization', 'get_organization', 'update_organization', 'get_org_branding', 'update_org_branding'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft organization operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _ORGANIZATION_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_organization":
            return await invoke_client_method(client.list_organization, **kwargs)
        if action == "get_organization":
            return await invoke_client_method(client.get_organization, **kwargs)
        if action == "update_organization":
            return await invoke_client_method(client.update_organization, **kwargs)
        if action == "get_org_branding":
            return await invoke_client_method(client.get_org_branding, **kwargs)
        if action == "update_org_branding":
            return await invoke_client_method(client.update_org_branding, **kwargs)
        raise ValueError(f"Unknown action: {action}")
