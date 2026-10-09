"""MCP tools for devices operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_connector_sdk.mcp.action_dispatch import resolve_action
from agent_connector_sdk.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_DEVICES_ACTIONS = (
    "list_devices",
    "get_device",
    "delete_device",
    "list_managed_devices",
    "get_managed_device",
    "list_device_compliance_policies",
    "list_device_configurations",
    "wipe_managed_device",
    "retire_managed_device",
)


def register_devices_tools(mcp: FastMCP):
    @mcp.tool(tags={"devices"})
    async def microsoft_devices(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_devices', 'get_device', 'delete_device', 'list_managed_devices', 'get_managed_device', 'list_device_compliance_policies', 'list_device_configurations', 'wipe_managed_device', 'retire_managed_device'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft devices operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _DEVICES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
