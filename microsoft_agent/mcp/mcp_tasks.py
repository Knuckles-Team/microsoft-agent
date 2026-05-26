"""MCP tools for tasks operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_tasks_tools(mcp: FastMCP):
    @mcp.tool(tags={"tasks"})
    async def microsoft_tasks(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_todo_task', 'create_todo_task', 'update_todo_task', 'delete_todo_task', 'get_planner_plan', 'get_planner_task', 'create_planner_task', 'update_planner_task', 'update_planner_task_details'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft tasks operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_todo_task":
            return client.get_todo_task(**kwargs)
        if action == "create_todo_task":
            return client.create_todo_task(**kwargs)
        if action == "update_todo_task":
            return client.update_todo_task(**kwargs)
        if action == "delete_todo_task":
            return client.delete_todo_task(**kwargs)
        if action == "get_planner_plan":
            return client.get_planner_plan(**kwargs)
        if action == "get_planner_task":
            return client.get_planner_task(**kwargs)
        if action == "create_planner_task":
            return client.create_planner_task(**kwargs)
        if action == "update_planner_task":
            return client.update_planner_task(**kwargs)
        if action == "update_planner_task_details":
            return client.update_planner_task_details(**kwargs)
        raise ValueError(f"Unknown action: {action}")
