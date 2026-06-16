"""MCP tools for identity operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_identity_tools(mcp: FastMCP):
    @mcp.tool(tags={"identity"})
    async def microsoft_identity(
        action: str = Field(
            description="Action to perform. Must be one of: 'create_invitation', 'list_conditional_access_policies', 'get_conditional_access_policy', 'create_conditional_access_policy', 'update_conditional_access_policy', 'delete_conditional_access_policy', 'list_access_reviews', 'get_access_review', 'list_entitlement_access_packages', 'list_lifecycle_workflows'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft identity operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "create_invitation":
            return await run_blocking(client.create_invitation, **kwargs)
        if action == "list_conditional_access_policies":
            return await run_blocking(client.list_conditional_access_policies, **kwargs)
        if action == "get_conditional_access_policy":
            return await run_blocking(client.get_conditional_access_policy, **kwargs)
        if action == "create_conditional_access_policy":
            return await run_blocking(client.create_conditional_access_policy, **kwargs)
        if action == "update_conditional_access_policy":
            return await run_blocking(client.update_conditional_access_policy, **kwargs)
        if action == "delete_conditional_access_policy":
            return await run_blocking(client.delete_conditional_access_policy, **kwargs)
        if action == "list_access_reviews":
            return await run_blocking(client.list_access_reviews, **kwargs)
        if action == "get_access_review":
            return await run_blocking(client.get_access_review, **kwargs)
        if action == "list_entitlement_access_packages":
            return await run_blocking(client.list_entitlement_access_packages, **kwargs)
        if action == "list_lifecycle_workflows":
            return await run_blocking(client.list_lifecycle_workflows, **kwargs)
        raise ValueError(f"Unknown action: {action}")
