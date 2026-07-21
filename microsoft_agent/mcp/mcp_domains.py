"""MCP tools for domains operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_DOMAINS_ACTIONS = (
    "list_domains",
    "get_domain",
    "create_domain",
    "delete_domain",
    "verify_domain",
    "list_domain_service_configuration_records",
)


def register_domains_tools(mcp: FastMCP):
    @mcp.tool(tags={"domains"})
    async def microsoft_domains(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_domains', 'get_domain', 'create_domain', 'delete_domain', 'verify_domain', 'list_domain_service_configuration_records'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
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
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _DOMAINS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_domains":
            return await invoke_client_method(client.list_domains, **kwargs)
        if action == "get_domain":
            return await invoke_client_method(client.get_domain, **kwargs)
        if action == "create_domain":
            return await invoke_client_method(client.create_domain, **kwargs)
        if action == "delete_domain":
            return await invoke_client_method(client.delete_domain, **kwargs)
        if action == "verify_domain":
            return await invoke_client_method(client.verify_domain, **kwargs)
        if action == "list_domain_service_configuration_records":
            return await invoke_client_method(
                client.list_domain_service_configuration_records, **kwargs
            )
        raise ValueError(f"Unknown action: {action}")
