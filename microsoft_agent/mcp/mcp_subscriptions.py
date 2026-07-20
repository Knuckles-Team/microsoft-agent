"""MCP tools for subscriptions operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_SUBSCRIPTIONS_ACTIONS = (
    "list_subscriptions",
    "get_subscription",
    "create_subscription",
    "update_subscription",
    "delete_subscription",
)


def register_subscriptions_tools(mcp: FastMCP):
    @mcp.tool(tags={"subscriptions"})
    async def microsoft_subscriptions(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_subscriptions', 'get_subscription', 'create_subscription', 'update_subscription', 'delete_subscription'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft subscriptions operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _SUBSCRIPTIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_subscriptions":
            return await invoke_client_method(client.list_subscriptions, **kwargs)
        if action == "get_subscription":
            return await invoke_client_method(client.get_subscription, **kwargs)
        if action == "create_subscription":
            return await invoke_client_method(client.create_subscription, **kwargs)
        if action == "update_subscription":
            return await invoke_client_method(client.update_subscription, **kwargs)
        if action == "delete_subscription":
            return await invoke_client_method(client.delete_subscription, **kwargs)
        raise ValueError(f"Unknown action: {action}")
