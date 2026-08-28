"""MCP tools for contacts operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_CONTACTS_ACTIONS = (
    "get_outlook_contact",
    "create_outlook_contact",
    "update_outlook_contact",
    "delete_outlook_contact",
)


def register_contacts_tools(mcp: FastMCP):
    @mcp.tool(tags={"contacts"})
    async def microsoft_contacts(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_outlook_contact', 'create_outlook_contact', 'update_outlook_contact', 'delete_outlook_contact'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft contacts operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _CONTACTS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_outlook_contact":
            return await invoke_client_method(client.get_outlook_contact, **kwargs)
        if action == "create_outlook_contact":
            return await invoke_client_method(client.create_outlook_contact, **kwargs)
        if action == "update_outlook_contact":
            return await invoke_client_method(client.update_outlook_contact, **kwargs)
        if action == "delete_outlook_contact":
            return await invoke_client_method(client.delete_outlook_contact, **kwargs)
        raise ValueError(f"Unknown action: {action}")
