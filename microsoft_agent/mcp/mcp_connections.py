"""MCP tools for connections operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_CONNECTIONS_ACTIONS = (
    "list_external_connections",
    "get_external_connection",
    "create_external_connection",
    "delete_external_connection",
)


def register_connections_tools(mcp: FastMCP):
    @mcp.tool(tags={"connections"})
    async def microsoft_connections(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_external_connections', 'get_external_connection', 'create_external_connection', 'delete_external_connection'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft connections operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _CONNECTIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_external_connections":
            return await invoke_client_method(
                client.list_external_connections, **kwargs
            )
        if action == "get_external_connection":
            return await invoke_client_method(client.get_external_connection, **kwargs)
        if action == "create_external_connection":
            return await invoke_client_method(
                client.create_external_connection, **kwargs
            )
        if action == "delete_external_connection":
            return await invoke_client_method(
                client.delete_external_connection, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
