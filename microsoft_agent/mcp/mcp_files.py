"""MCP tools for files operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_files_tools(mcp: FastMCP):
    @mcp.tool(tags={"files"})
    async def microsoft_files(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_users', 'list_drives', 'get_drive_root_item', 'download_onedrive_file_content', 'delete_onedrive_file', 'upload_file_content', 'create_excel_chart', 'format_excel_range', 'sort_excel_range', 'get_excel_range', 'list_excel_worksheets', 'list_excel_tables', 'get_excel_workbook', 'list_onenote_notebooks', 'list_onenote_notebook_sections', 'list_onenote_section_pages', 'list_todo_task_lists', 'list_todo_tasks', 'list_planner_tasks', 'list_plan_tasks', 'list_outlook_contacts', 'list_chats', 'get_excel_worksheet', 'list_joined_teams', 'list_team_channels', 'list_team_members', 'list_site_drives', 'get_site_drive_by_id', 'list_site_items', 'get_site_item', 'list_site_lists', 'get_site_list', 'list_sharepoint_site_list_items', 'get_sharepoint_site_list_item', 'get_excel_table'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft files operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_users":
            return await run_blocking(client.list_users, **kwargs)
        if action == "list_drives":
            return await run_blocking(client.list_drives, **kwargs)
        if action == "get_drive_root_item":
            return await run_blocking(client.get_drive_root_item, **kwargs)
        if action == "download_onedrive_file_content":
            return await run_blocking(client.download_onedrive_file_content, **kwargs)
        if action == "delete_onedrive_file":
            return await run_blocking(client.delete_onedrive_file, **kwargs)
        if action == "upload_file_content":
            return await run_blocking(client.upload_file_content, **kwargs)
        if action == "create_excel_chart":
            return await run_blocking(client.create_excel_chart, **kwargs)
        if action == "format_excel_range":
            return await run_blocking(client.format_excel_range, **kwargs)
        if action == "sort_excel_range":
            return await run_blocking(client.sort_excel_range, **kwargs)
        if action == "get_excel_range":
            return await run_blocking(client.get_excel_range, **kwargs)
        if action == "list_excel_worksheets":
            return await run_blocking(client.list_excel_worksheets, **kwargs)
        if action == "list_excel_tables":
            return await run_blocking(client.list_excel_tables, **kwargs)
        if action == "get_excel_workbook":
            return await run_blocking(client.get_excel_workbook, **kwargs)
        if action == "list_onenote_notebooks":
            return await run_blocking(client.list_onenote_notebooks, **kwargs)
        if action == "list_onenote_notebook_sections":
            return await run_blocking(client.list_onenote_notebook_sections, **kwargs)
        if action == "list_onenote_section_pages":
            return await run_blocking(client.list_onenote_section_pages, **kwargs)
        if action == "list_todo_task_lists":
            return await run_blocking(client.list_todo_task_lists, **kwargs)
        if action == "list_todo_tasks":
            return await run_blocking(client.list_todo_tasks, **kwargs)
        if action == "list_planner_tasks":
            return await run_blocking(client.list_planner_tasks, **kwargs)
        if action == "list_plan_tasks":
            return await run_blocking(client.list_plan_tasks, **kwargs)
        if action == "list_outlook_contacts":
            return await run_blocking(client.list_outlook_contacts, **kwargs)
        if action == "list_chats":
            return await run_blocking(client.list_chats, **kwargs)
        if action == "get_excel_worksheet":
            return await run_blocking(client.get_excel_worksheet, **kwargs)
        if action == "list_joined_teams":
            return await run_blocking(client.list_joined_teams, **kwargs)
        if action == "list_team_channels":
            return await run_blocking(client.list_team_channels, **kwargs)
        if action == "list_team_members":
            return await run_blocking(client.list_team_members, **kwargs)
        if action == "list_site_drives":
            return await run_blocking(client.list_site_drives, **kwargs)
        if action == "get_site_drive_by_id":
            return await run_blocking(client.get_site_drive_by_id, **kwargs)
        if action == "list_site_items":
            return await run_blocking(client.list_site_items, **kwargs)
        if action == "get_site_item":
            return await run_blocking(client.get_site_item, **kwargs)
        if action == "list_site_lists":
            return await run_blocking(client.list_site_lists, **kwargs)
        if action == "get_site_list":
            return await run_blocking(client.get_site_list, **kwargs)
        if action == "list_sharepoint_site_list_items":
            return await run_blocking(client.list_sharepoint_site_list_items, **kwargs)
        if action == "get_sharepoint_site_list_item":
            return await run_blocking(client.get_sharepoint_site_list_item, **kwargs)
        if action == "get_excel_table":
            return await run_blocking(client.get_excel_table, **kwargs)
        raise ValueError(f"Unknown action: {action}")
