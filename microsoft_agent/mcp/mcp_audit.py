"""MCP tools for audit operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_audit_tools(mcp: FastMCP):
    @mcp.tool(tags={"audit"})
    async def microsoft_audit(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_directory_audits', 'get_directory_audit', 'list_sign_in_logs', 'get_sign_in_log', 'list_provisioning_logs'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
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
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_directory_audits":
            return client.list_directory_audits(**kwargs)
        if action == "get_directory_audit":
            return client.get_directory_audit(**kwargs)
        if action == "list_sign_in_logs":
            return client.list_sign_in_logs(**kwargs)
        if action == "get_sign_in_log":
            return client.get_sign_in_log(**kwargs)
        if action == "list_provisioning_logs":
            return client.list_provisioning_logs(**kwargs)
        raise ValueError(f"Unknown action: {action}")
