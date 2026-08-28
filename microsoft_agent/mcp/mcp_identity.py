"""MCP tools for identity operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_IDENTITY_ACTIONS = (
    "create_invitation",
    "list_conditional_access_policies",
    "get_conditional_access_policy",
    "create_conditional_access_policy",
    "update_conditional_access_policy",
    "delete_conditional_access_policy",
    "list_access_reviews",
    "get_access_review",
    "list_entitlement_access_packages",
    "list_lifecycle_workflows",
)


def register_identity_tools(mcp: FastMCP):
    @mcp.tool(tags={"identity"})
    async def microsoft_identity(
        action: str = Field(
            description="Action to perform. Must be one of: 'create_invitation', 'list_conditional_access_policies', 'get_conditional_access_policy', 'create_conditional_access_policy', 'update_conditional_access_policy', 'delete_conditional_access_policy', 'list_access_reviews', 'get_access_review', 'list_entitlement_access_packages', 'list_lifecycle_workflows'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft identity operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _IDENTITY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
