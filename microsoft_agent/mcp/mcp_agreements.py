"""MCP tools for agreements operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_AGREEMENTS_ACTIONS = (
    "list_agreements",
    "get_agreement",
    "create_agreement",
    "delete_agreement",
)


def register_agreements_tools(mcp: FastMCP):
    @mcp.tool(tags={"agreements"})
    async def microsoft_agreements(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_agreements', 'get_agreement', 'create_agreement', 'delete_agreement'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft agreements operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _AGREEMENTS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_agreements":
            return await invoke_client_method(client.list_agreements, **kwargs)
        if action == "get_agreement":
            return await invoke_client_method(client.get_agreement, **kwargs)
        if action == "create_agreement":
            return await invoke_client_method(client.create_agreement, **kwargs)
        if action == "delete_agreement":
            return await invoke_client_method(client.delete_agreement, **kwargs)
        raise ValueError(f"Unknown action: {action}")
