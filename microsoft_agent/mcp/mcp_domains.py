"""MCP tools for domains operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_domains_tools(mcp: FastMCP):
    @mcp.tool(tags={"domains"})
    async def microsoft_domains(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_domains', 'get_domain', 'create_domain', 'delete_domain', 'verify_domain', 'list_domain_service_configuration_records'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft domains operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_domains":
            return client.list_domains(**kwargs)
        if action == "get_domain":
            return client.get_domain(**kwargs)
        if action == "create_domain":
            return client.create_domain(**kwargs)
        if action == "delete_domain":
            return client.delete_domain(**kwargs)
        if action == "verify_domain":
            return client.verify_domain(**kwargs)
        if action == "list_domain_service_configuration_records":
            return client.list_domain_service_configuration_records(**kwargs)
        raise ValueError(f"Unknown action: {action}")
