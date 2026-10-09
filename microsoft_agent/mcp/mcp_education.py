"""MCP tools for education operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_connector_sdk.mcp.action_dispatch import resolve_action
from agent_connector_sdk.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_EDUCATION_ACTIONS = (
    "list_education_classes",
    "get_education_class",
    "list_education_schools",
    "get_education_school",
    "list_education_users",
    "list_education_assignments",
)


def register_education_tools(mcp: FastMCP):
    @mcp.tool(tags={"education"})
    async def microsoft_education(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_education_classes', 'get_education_class', 'list_education_schools', 'get_education_school', 'list_education_users', 'list_education_assignments'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft education operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _EDUCATION_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
