"""MCP tools for applications operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_applications_tools(mcp: FastMCP):
    @mcp.tool(tags={"applications"})
    async def microsoft_applications(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_applications', 'get_application', 'create_application', 'update_application', 'delete_application', 'add_application_password', 'remove_application_password', 'list_service_principals', 'get_service_principal', 'create_service_principal', 'update_service_principal', 'delete_service_principal'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft applications operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_applications":
            return await run_blocking(client.list_applications, **kwargs)
        if action == "get_application":
            return await run_blocking(client.get_application, **kwargs)
        if action == "create_application":
            return await run_blocking(client.create_application, **kwargs)
        if action == "update_application":
            return await run_blocking(client.update_application, **kwargs)
        if action == "delete_application":
            return await run_blocking(client.delete_application, **kwargs)
        if action == "add_application_password":
            return await run_blocking(client.add_application_password, **kwargs)
        if action == "remove_application_password":
            return await run_blocking(client.remove_application_password, **kwargs)
        if action == "list_service_principals":
            return await run_blocking(client.list_service_principals, **kwargs)
        if action == "get_service_principal":
            return await run_blocking(client.get_service_principal, **kwargs)
        if action == "create_service_principal":
            return await run_blocking(client.create_service_principal, **kwargs)
        if action == "update_service_principal":
            return await run_blocking(client.update_service_principal, **kwargs)
        if action == "delete_service_principal":
            return await run_blocking(client.delete_service_principal, **kwargs)
        raise ValueError(f"Unknown action: {action}")
