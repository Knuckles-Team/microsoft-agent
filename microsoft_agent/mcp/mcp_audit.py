"""MCP tools for audit operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_AUDIT_ACTIONS = (
    "list_directory_audits",
    "get_directory_audit",
    "list_sign_in_logs",
    "get_sign_in_log",
    "list_provisioning_logs",
)


def register_audit_tools(mcp: FastMCP):
    @mcp.tool(tags={"audit"})
    async def microsoft_audit(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_directory_audits', 'get_directory_audit', 'list_sign_in_logs', 'get_sign_in_log', 'list_provisioning_logs'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft audit operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _AUDIT_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_directory_audits":
            return await invoke_client_method(client.list_directory_audits, **kwargs)
        if action == "get_directory_audit":
            return await invoke_client_method(client.get_directory_audit, **kwargs)
        if action == "list_sign_in_logs":
            return await invoke_client_method(client.list_sign_in_logs, **kwargs)
        if action == "get_sign_in_log":
            return await invoke_client_method(client.get_sign_in_log, **kwargs)
        if action == "list_provisioning_logs":
            return await invoke_client_method(client.list_provisioning_logs, **kwargs)
        raise ValueError(f"Unknown action: {action}")
