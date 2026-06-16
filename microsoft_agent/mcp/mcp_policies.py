"""MCP tools for policies operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_policies_tools(mcp: FastMCP):
    @mcp.tool(tags={"policies"})
    async def microsoft_policies(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_authorization_policy', 'list_token_lifetime_policies', 'list_token_issuance_policies', 'list_permission_grant_policies', 'get_admin_consent_policy'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft policies operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_authorization_policy":
            return await run_blocking(client.get_authorization_policy, **kwargs)
        if action == "list_token_lifetime_policies":
            return await run_blocking(client.list_token_lifetime_policies, **kwargs)
        if action == "list_token_issuance_policies":
            return await run_blocking(client.list_token_issuance_policies, **kwargs)
        if action == "list_permission_grant_policies":
            return await run_blocking(client.list_permission_grant_policies, **kwargs)
        if action == "get_admin_consent_policy":
            return await run_blocking(client.get_admin_consent_policy, **kwargs)
        raise ValueError(f"Unknown action: {action}")
