"""MCP tools for devices operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_devices_tools(mcp: FastMCP):
    @mcp.tool(tags={"devices"})
    async def microsoft_devices(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_devices', 'get_device', 'delete_device', 'list_managed_devices', 'get_managed_device', 'list_device_compliance_policies', 'list_device_configurations', 'wipe_managed_device', 'retire_managed_device'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft devices operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_devices":
            return client.list_devices(**kwargs)
        if action == "get_device":
            return client.get_device(**kwargs)
        if action == "delete_device":
            return client.delete_device(**kwargs)
        if action == "list_managed_devices":
            return client.list_managed_devices(**kwargs)
        if action == "get_managed_device":
            return client.get_managed_device(**kwargs)
        if action == "list_device_compliance_policies":
            return client.list_device_compliance_policies(**kwargs)
        if action == "list_device_configurations":
            return client.list_device_configurations(**kwargs)
        if action == "wipe_managed_device":
            return client.wipe_managed_device(**kwargs)
        if action == "retire_managed_device":
            return client.retire_managed_device(**kwargs)
        raise ValueError(f"Unknown action: {action}")
