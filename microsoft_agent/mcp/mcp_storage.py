"""MCP tools for storage operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_STORAGE_ACTIONS = (
    "list_file_storage_containers",
    "get_file_storage_container",
    "create_file_storage_container",
)


def register_storage_tools(mcp: FastMCP):
    @mcp.tool(tags={"storage"})
    async def microsoft_storage(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_file_storage_containers', 'get_file_storage_container', 'create_file_storage_container'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft storage operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _STORAGE_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_file_storage_containers":
            return await invoke_client_method(
                client.list_file_storage_containers, **kwargs
            )
        if action == "get_file_storage_container":
            return await invoke_client_method(
                client.get_file_storage_container, **kwargs
            )
        if action == "create_file_storage_container":
            return await invoke_client_method(
                client.create_file_storage_container, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
