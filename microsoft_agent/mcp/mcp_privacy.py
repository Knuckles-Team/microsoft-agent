"""MCP tools for privacy operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_PRIVACY_ACTIONS = (
    "list_subject_rights_requests",
    "get_subject_rights_request",
    "create_subject_rights_request",
)


def register_privacy_tools(mcp: FastMCP):
    @mcp.tool(tags={"privacy"})
    async def microsoft_privacy(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_subject_rights_requests', 'get_subject_rights_request', 'create_subject_rights_request'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft privacy operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _PRIVACY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_subject_rights_requests":
            return await invoke_client_method(
                client.list_subject_rights_requests, **kwargs
            )
        if action == "get_subject_rights_request":
            return await invoke_client_method(
                client.get_subject_rights_request, **kwargs
            )
        if action == "create_subject_rights_request":
            return await invoke_client_method(
                client.create_subject_rights_request, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
