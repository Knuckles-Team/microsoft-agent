"""MCP tools for auth operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_AUTH_ACTIONS = ("login", "logout", "verify_login", "list_accounts")


def register_auth_tools(mcp: FastMCP):
    @mcp.tool(tags={"auth"})
    async def microsoft_auth(
        action: str = Field(
            description="Action to perform. Must be one of: 'login', 'logout', 'verify_login', 'list_accounts'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft auth operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _AUTH_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "login":
            return await invoke_client_method(client.login, **kwargs)
        if action == "logout":
            return await invoke_client_method(client.logout, **kwargs)
        if action == "verify_login":
            return await invoke_client_method(client.verify_login, **kwargs)
        if action == "list_accounts":
            return await invoke_client_method(client.list_accounts, **kwargs)
        raise ValueError(f"Unknown action: {action}")
