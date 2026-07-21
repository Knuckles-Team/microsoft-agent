"""MCP tools for communications operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_COMMUNICATIONS_ACTIONS = (
    "list_online_meetings",
    "get_online_meeting",
    "create_online_meeting",
    "update_online_meeting",
    "delete_online_meeting",
    "list_call_records",
    "get_call_record",
    "list_presences",
    "get_presence",
    "get_my_presence",
)


def register_communications_tools(mcp: FastMCP):
    @mcp.tool(tags={"communications"})
    async def microsoft_communications(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_online_meetings', 'get_online_meeting', 'create_online_meeting', 'update_online_meeting', 'delete_online_meeting', 'list_call_records', 'get_call_record', 'list_presences', 'get_presence', 'get_my_presence'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft communications operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(
            action, _COMMUNICATIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_online_meetings":
            return await invoke_client_method(client.list_online_meetings, **kwargs)
        if action == "get_online_meeting":
            return await invoke_client_method(client.get_online_meeting, **kwargs)
        if action == "create_online_meeting":
            return await invoke_client_method(client.create_online_meeting, **kwargs)
        if action == "update_online_meeting":
            return await invoke_client_method(client.update_online_meeting, **kwargs)
        if action == "delete_online_meeting":
            return await invoke_client_method(client.delete_online_meeting, **kwargs)
        if action == "list_call_records":
            return await invoke_client_method(client.list_call_records, **kwargs)
        if action == "get_call_record":
            return await invoke_client_method(client.get_call_record, **kwargs)
        if action == "list_presences":
            return await invoke_client_method(client.list_presences, **kwargs)
        if action == "get_presence":
            return await invoke_client_method(client.get_presence, **kwargs)
        if action == "get_my_presence":
            return await invoke_client_method(client.get_my_presence, **kwargs)
        raise ValueError(f"Unknown action: {action}")
