"""MCP tools for print operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_connector_sdk.mcp.action_dispatch import resolve_action
from agent_connector_sdk.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_PRINT_ACTIONS = (
    "list_printers",
    "get_printer",
    "list_print_jobs",
    "create_print_job",
    "create_print_document_upload_session",
    "start_print_job",
    "submit_print_document",
    "list_print_shares",
)


def register_print_tools(mcp: FastMCP):
    @mcp.tool(tags={"print"})
    async def microsoft_print(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_printers', 'get_printer', 'list_print_jobs', 'create_print_job', 'create_print_document_upload_session', 'start_print_job', 'submit_print_document', 'list_print_shares'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft print operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _PRINT_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
