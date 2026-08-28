"""MCP tools for files operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_FILES_ACTIONS = (
    "list_users",
    "list_drives",
    "get_drive_root_item",
    "download_onedrive_file_content",
    "delete_onedrive_file",
    "upload_file_content",
    "create_excel_chart",
    "format_excel_range",
    "sort_excel_range",
    "get_excel_range",
    "list_excel_worksheets",
    "list_excel_tables",
    "get_excel_workbook",
    "list_onenote_notebooks",
    "list_onenote_notebook_sections",
    "list_onenote_section_pages",
    "list_todo_task_lists",
    "list_todo_tasks",
    "list_planner_tasks",
    "list_plan_tasks",
    "list_outlook_contacts",
    "list_chats",
    "get_excel_worksheet",
    "list_joined_teams",
    "list_team_channels",
    "list_team_members",
    "list_site_drives",
    "get_site_drive_by_id",
    "list_site_items",
    "get_site_item",
    "list_site_lists",
    "get_site_list",
    "list_sharepoint_site_list_items",
    "get_sharepoint_site_list_item",
    "get_excel_table",
)


def register_files_tools(mcp: FastMCP):
    @mcp.tool(tags={"files"})
    async def microsoft_files(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_users', 'list_drives', 'get_drive_root_item', 'download_onedrive_file_content', 'delete_onedrive_file', 'upload_file_content', 'create_excel_chart', 'format_excel_range', 'sort_excel_range', 'get_excel_range', 'list_excel_worksheets', 'list_excel_tables', 'get_excel_workbook', 'list_onenote_notebooks', 'list_onenote_notebook_sections', 'list_onenote_section_pages', 'list_todo_task_lists', 'list_todo_tasks', 'list_planner_tasks', 'list_plan_tasks', 'list_outlook_contacts', 'list_chats', 'get_excel_worksheet', 'list_joined_teams', 'list_team_channels', 'list_team_members', 'list_site_drives', 'get_site_drive_by_id', 'list_site_items', 'get_site_item', 'list_site_lists', 'get_site_list', 'list_sharepoint_site_list_items', 'get_sharepoint_site_list_item', 'get_excel_table'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft files operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _FILES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
