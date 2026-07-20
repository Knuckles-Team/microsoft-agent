"""MCP tools for print operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
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
            ctx.info("Executing tool...")
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

        if action == "list_printers":
            return await invoke_client_method(client.list_printers, **kwargs)
        if action == "get_printer":
            return await invoke_client_method(client.get_printer, **kwargs)
        if action == "list_print_jobs":
            return await invoke_client_method(client.list_print_jobs, **kwargs)
        if action == "create_print_job":
            return await invoke_client_method(client.create_print_job, **kwargs)
        if action == "create_print_document_upload_session":
            return await invoke_client_method(
                client.create_print_document_upload_session, **kwargs
            )
        if action == "start_print_job":
            return await invoke_client_method(client.start_print_job, **kwargs)
        if action == "submit_print_document":
            return await invoke_client_method(client.submit_print_document, **kwargs)
        if action == "list_print_shares":
            return await invoke_client_method(client.list_print_shares, **kwargs)
        raise ValueError(f"Unknown action: {action}")
