"""MCP tools for employee experience operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_EMPLOYEE_EXPERIENCE_ACTIONS = (
    "list_learning_providers",
    "get_learning_provider",
    "list_learning_course_activities",
)


def register_employee_experience_tools(mcp: FastMCP):
    @mcp.tool(tags={"employee_experience"})
    async def microsoft_employee_experience(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_learning_providers', 'get_learning_provider', 'list_learning_course_activities'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft employee experience operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _EMPLOYEE_EXPERIENCE_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_learning_providers":
            return await invoke_client_method(client.list_learning_providers, **kwargs)
        if action == "get_learning_provider":
            return await invoke_client_method(client.get_learning_provider, **kwargs)
        if action == "list_learning_course_activities":
            return await invoke_client_method(
                client.list_learning_course_activities, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
