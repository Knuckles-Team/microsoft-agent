"""MCP tools for tasks operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_TASKS_ACTIONS = (
    "get_todo_task",
    "create_todo_task",
    "update_todo_task",
    "delete_todo_task",
    "get_planner_plan",
    "get_planner_task",
    "create_planner_task",
    "update_planner_task",
    "update_planner_task_details",
)


def register_tasks_tools(mcp: FastMCP):
    @mcp.tool(tags={"tasks"})
    async def microsoft_tasks(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_todo_task', 'create_todo_task', 'update_todo_task', 'delete_todo_task', 'get_planner_plan', 'get_planner_task', 'create_planner_task', 'update_planner_task', 'update_planner_task_details'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft tasks operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _TASKS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_todo_task":
            return await invoke_client_method(client.get_todo_task, **kwargs)
        if action == "create_todo_task":
            return await invoke_client_method(client.create_todo_task, **kwargs)
        if action == "update_todo_task":
            return await invoke_client_method(client.update_todo_task, **kwargs)
        if action == "delete_todo_task":
            return await invoke_client_method(client.delete_todo_task, **kwargs)
        if action == "get_planner_plan":
            return await invoke_client_method(client.get_planner_plan, **kwargs)
        if action == "get_planner_task":
            return await invoke_client_method(client.get_planner_task, **kwargs)
        if action == "create_planner_task":
            return await invoke_client_method(client.create_planner_task, **kwargs)
        if action == "update_planner_task":
            return await invoke_client_method(client.update_planner_task, **kwargs)
        if action == "update_planner_task_details":
            return await invoke_client_method(
                client.update_planner_task_details, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
