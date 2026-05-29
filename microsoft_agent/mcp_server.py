#!/usr/bin/python
"""
Microsoft Graph MCP Server implementation.
"""

import warnings

from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from fastmcp.utilities.logging import get_logger
from pydantic import Field

# Filter RequestsDependencyWarning early to prevent log spam
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    try:
        from requests.exceptions import RequestsDependencyWarning

        warnings.filterwarnings("ignore", category=RequestsDependencyWarning)
    except ImportError:
        pass

warnings.filterwarnings("ignore", message=".*urllib3.*or chardet.*")
warnings.filterwarnings("ignore", message=".*urllib3.*or charset_normalizer.*")

import logging
import os
import sys
from typing import Any

from agent_utilities.base_utilities import to_boolean
from agent_utilities.mcp_utilities import create_mcp_server
from dotenv import find_dotenv, load_dotenv
from starlette.requests import Request
from starlette.responses import JSONResponse

from microsoft_agent.auth import get_client

__version__ = "0.22.0"

logger = get_logger(name="microsoft-agent")
logger.setLevel(logging.INFO)


def register_auth_tools(mcp: FastMCP):
    @mcp.tool(tags={"auth"})
    async def microsoft_auth(
        action: str = Field(
            description="Action to perform. Must be one of: 'login', 'logout', 'verify_login', 'list_accounts'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft auth operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "login":
            return client.login(**kwargs)
        if action == "logout":
            return client.logout(**kwargs)
        if action == "verify_login":
            return client.verify_login(**kwargs)
        if action == "list_accounts":
            return client.list_accounts(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_meta_tools(mcp: FastMCP):
    @mcp.tool(tags={"meta"})
    async def microsoft_meta(
        action: str = Field(
            description="Action to perform. Must be one of: 'searches'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft meta operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "searches":
            return client.searches(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_mail_tools(mcp: FastMCP):
    @mcp.tool(tags={"mail"})
    async def microsoft_mail(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_mail_messages', 'list_mail_folders', 'list_mail_folder_messages', 'get_mail_message', 'send_mail', 'list_shared_mailbox_messages', 'list_shared_mailbox_folder_messages', 'get_shared_mailbox_message', 'send_shared_mailbox_mail', 'create_draft_email', 'delete_mail_message', 'move_mail_message', 'update_mail_message', 'add_mail_attachment', 'list_mail_attachments', 'get_mail_attachment', 'delete_mail_attachment', 'get_root_folder', 'list_folder_files', 'list_chat_messages', 'get_chat_message', 'send_chat_message', 'list_channel_messages', 'get_channel_message', 'send_channel_message', 'list_chat_message_replies', 'reply_to_chat_message'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft mail operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_mail_messages":
            return client.list_mail_messages(**kwargs)
        if action == "list_mail_folders":
            return client.list_mail_folders(**kwargs)
        if action == "list_mail_folder_messages":
            return client.list_mail_folder_messages(**kwargs)
        if action == "get_mail_message":
            return client.get_mail_message(**kwargs)
        if action == "send_mail":
            return client.send_mail(**kwargs)
        if action == "list_shared_mailbox_messages":
            return client.list_shared_mailbox_messages(**kwargs)
        if action == "list_shared_mailbox_folder_messages":
            return client.list_shared_mailbox_folder_messages(**kwargs)
        if action == "get_shared_mailbox_message":
            return client.get_shared_mailbox_message(**kwargs)
        if action == "send_shared_mailbox_mail":
            return client.send_shared_mailbox_mail(**kwargs)
        if action == "create_draft_email":
            return client.create_draft_email(**kwargs)
        if action == "delete_mail_message":
            return client.delete_mail_message(**kwargs)
        if action == "move_mail_message":
            return client.move_mail_message(**kwargs)
        if action == "update_mail_message":
            return client.update_mail_message(**kwargs)
        if action == "add_mail_attachment":
            return client.add_mail_attachment(**kwargs)
        if action == "list_mail_attachments":
            return client.list_mail_attachments(**kwargs)
        if action == "get_mail_attachment":
            return client.get_mail_attachment(**kwargs)
        if action == "delete_mail_attachment":
            return client.delete_mail_attachment(**kwargs)
        if action == "get_root_folder":
            return client.get_root_folder(**kwargs)
        if action == "list_folder_files":
            return client.list_folder_files(**kwargs)
        if action == "list_chat_messages":
            return client.list_chat_messages(**kwargs)
        if action == "get_chat_message":
            return client.get_chat_message(**kwargs)
        if action == "send_chat_message":
            return client.send_chat_message(**kwargs)
        if action == "list_channel_messages":
            return client.list_channel_messages(**kwargs)
        if action == "get_channel_message":
            return client.get_channel_message(**kwargs)
        if action == "send_channel_message":
            return client.send_channel_message(**kwargs)
        if action == "list_chat_message_replies":
            return client.list_chat_message_replies(**kwargs)
        if action == "reply_to_chat_message":
            return client.reply_to_chat_message(**kwargs)
        raise ValueError(f"Unknown action: {action}")


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
            return client.list_users(**kwargs)
        if action == "list_drives":
            return client.list_drives(**kwargs)
        if action == "get_drive_root_item":
            return client.get_drive_root_item(**kwargs)
        if action == "download_onedrive_file_content":
            return client.download_onedrive_file_content(**kwargs)
        if action == "delete_onedrive_file":
            return client.delete_onedrive_file(**kwargs)
        if action == "upload_file_content":
            return client.upload_file_content(**kwargs)
        if action == "create_excel_chart":
            return client.create_excel_chart(**kwargs)
        if action == "format_excel_range":
            return client.format_excel_range(**kwargs)
        if action == "sort_excel_range":
            return client.sort_excel_range(**kwargs)
        if action == "get_excel_range":
            return client.get_excel_range(**kwargs)
        if action == "list_excel_worksheets":
            return client.list_excel_worksheets(**kwargs)
        if action == "list_excel_tables":
            return client.list_excel_tables(**kwargs)
        if action == "get_excel_workbook":
            return client.get_excel_workbook(**kwargs)
        if action == "list_onenote_notebooks":
            return client.list_onenote_notebooks(**kwargs)
        if action == "list_onenote_notebook_sections":
            return client.list_onenote_notebook_sections(**kwargs)
        if action == "list_onenote_section_pages":
            return client.list_onenote_section_pages(**kwargs)
        if action == "list_todo_task_lists":
            return client.list_todo_task_lists(**kwargs)
        if action == "list_todo_tasks":
            return client.list_todo_tasks(**kwargs)
        if action == "list_planner_tasks":
            return client.list_planner_tasks(**kwargs)
        if action == "list_plan_tasks":
            return client.list_plan_tasks(**kwargs)
        if action == "list_outlook_contacts":
            return client.list_outlook_contacts(**kwargs)
        if action == "list_chats":
            return client.list_chats(**kwargs)
        if action == "get_excel_worksheet":
            return client.get_excel_worksheet(**kwargs)
        if action == "list_joined_teams":
            return client.list_joined_teams(**kwargs)
        if action == "list_team_channels":
            return client.list_team_channels(**kwargs)
        if action == "list_team_members":
            return client.list_team_members(**kwargs)
        if action == "list_site_drives":
            return client.list_site_drives(**kwargs)
        if action == "get_site_drive_by_id":
            return client.get_site_drive_by_id(**kwargs)
        if action == "list_site_items":
            return client.list_site_items(**kwargs)
        if action == "get_site_item":
            return client.get_site_item(**kwargs)
        if action == "list_site_lists":
            return client.list_site_lists(**kwargs)
        if action == "get_site_list":
            return client.get_site_list(**kwargs)
        if action == "list_sharepoint_site_list_items":
            return client.list_sharepoint_site_list_items(**kwargs)
        if action == "get_sharepoint_site_list_item":
            return client.get_sharepoint_site_list_item(**kwargs)
        if action == "get_excel_table":
            return client.get_excel_table(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_calendar_tools(mcp: FastMCP):
    @mcp.tool(tags={"calendar"})
    async def microsoft_calendar(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_calendar_events', 'get_calendar_event', 'create_calendar_event', 'update_calendar_event', 'delete_calendar_event', 'list_specific_calendar_events', 'get_specific_calendar_event', 'create_specific_calendar_event', 'update_specific_calendar_event', 'delete_specific_calendar_event', 'get_calendar_view', 'list_calendars', 'find_meeting_times'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft calendar operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_calendar_events":
            return client.list_calendar_events(**kwargs)
        if action == "get_calendar_event":
            return client.get_calendar_event(**kwargs)
        if action == "create_calendar_event":
            return client.create_calendar_event(**kwargs)
        if action == "update_calendar_event":
            return client.update_calendar_event(**kwargs)
        if action == "delete_calendar_event":
            return client.delete_calendar_event(**kwargs)
        if action == "list_specific_calendar_events":
            return client.list_specific_calendar_events(**kwargs)
        if action == "get_specific_calendar_event":
            return client.get_specific_calendar_event(**kwargs)
        if action == "create_specific_calendar_event":
            return client.create_specific_calendar_event(**kwargs)
        if action == "update_specific_calendar_event":
            return client.update_specific_calendar_event(**kwargs)
        if action == "delete_specific_calendar_event":
            return client.delete_specific_calendar_event(**kwargs)
        if action == "get_calendar_view":
            return client.get_calendar_view(**kwargs)
        if action == "list_calendars":
            return client.list_calendars(**kwargs)
        if action == "find_meeting_times":
            return client.find_meeting_times(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_notes_tools(mcp: FastMCP):
    @mcp.tool(tags={"notes"})
    async def microsoft_notes(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_onenote_page_content', 'create_onenote_page'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft notes operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_onenote_page_content":
            return client.get_onenote_page_content(**kwargs)
        if action == "create_onenote_page":
            return client.create_onenote_page(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_tasks_tools(mcp: FastMCP):
    @mcp.tool(tags={"tasks"})
    async def microsoft_tasks(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_todo_task', 'create_todo_task', 'update_todo_task', 'delete_todo_task', 'get_planner_plan', 'get_planner_task', 'create_planner_task', 'update_planner_task', 'update_planner_task_details'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft tasks operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_todo_task":
            return client.get_todo_task(**kwargs)
        if action == "create_todo_task":
            return client.create_todo_task(**kwargs)
        if action == "update_todo_task":
            return client.update_todo_task(**kwargs)
        if action == "delete_todo_task":
            return client.delete_todo_task(**kwargs)
        if action == "get_planner_plan":
            return client.get_planner_plan(**kwargs)
        if action == "get_planner_task":
            return client.get_planner_task(**kwargs)
        if action == "create_planner_task":
            return client.create_planner_task(**kwargs)
        if action == "update_planner_task":
            return client.update_planner_task(**kwargs)
        if action == "update_planner_task_details":
            return client.update_planner_task_details(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_contacts_tools(mcp: FastMCP):
    @mcp.tool(tags={"contacts"})
    async def microsoft_contacts(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_outlook_contact', 'create_outlook_contact', 'update_outlook_contact', 'delete_outlook_contact'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft contacts operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_outlook_contact":
            return client.get_outlook_contact(**kwargs)
        if action == "create_outlook_contact":
            return client.create_outlook_contact(**kwargs)
        if action == "update_outlook_contact":
            return client.update_outlook_contact(**kwargs)
        if action == "delete_outlook_contact":
            return client.delete_outlook_contact(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_user_tools(mcp: FastMCP):
    @mcp.tool(tags={"user"})
    async def microsoft_user(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_current_user', 'get_me'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft user operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_current_user":
            return client.get_current_user(**kwargs)
        if action == "get_me":
            return client.get_me(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_chat_tools(mcp: FastMCP):
    @mcp.tool(tags={"chat"})
    async def microsoft_chat(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_chat'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft chat operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_chat":
            return client.get_chat(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_teams_tools(mcp: FastMCP):
    @mcp.tool(tags={"teams"})
    async def microsoft_teams(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_team', 'get_team_channel'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft teams operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_team":
            return client.get_team(**kwargs)
        if action == "get_team_channel":
            return client.get_team_channel(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_sites_tools(mcp: FastMCP):
    @mcp.tool(tags={"sites"})
    async def microsoft_sites(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_sites', 'get_site', 'get_sharepoint_site_by_path', 'get_sharepoint_sites_delta'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft sites operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_sites":
            return client.list_sites(**kwargs)
        if action == "get_site":
            return client.get_site(**kwargs)
        if action == "get_sharepoint_site_by_path":
            return client.get_sharepoint_site_by_path(**kwargs)
        if action == "get_sharepoint_sites_delta":
            return client.get_sharepoint_sites_delta(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_search_tools(mcp: FastMCP):
    @mcp.tool(tags={"search"})
    async def microsoft_search(
        action: str = Field(
            description="Action to perform. Must be one of: 'search_query', 'search_tools'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft search operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "search_query":
            return client.search_query(**kwargs)
        if action == "search_tools":
            return client.search_tools(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_groups_tools(mcp: FastMCP):
    @mcp.tool(tags={"groups"})
    async def microsoft_groups(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_groups', 'get_group', 'create_group', 'update_group', 'delete_group', 'list_group_members', 'add_group_member', 'remove_group_member', 'list_group_owners', 'list_group_conversations', 'list_group_drives'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft groups operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_groups":
            return client.list_groups(**kwargs)
        if action == "get_group":
            return client.get_group(**kwargs)
        if action == "create_group":
            return client.create_group(**kwargs)
        if action == "update_group":
            return client.update_group(**kwargs)
        if action == "delete_group":
            return client.delete_group(**kwargs)
        if action == "list_group_members":
            return client.list_group_members(**kwargs)
        if action == "add_group_member":
            return client.add_group_member(**kwargs)
        if action == "remove_group_member":
            return client.remove_group_member(**kwargs)
        if action == "list_group_owners":
            return client.list_group_owners(**kwargs)
        if action == "list_group_conversations":
            return client.list_group_conversations(**kwargs)
        if action == "list_group_drives":
            return client.list_group_drives(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_admin_tools(mcp: FastMCP):
    @mcp.tool(tags={"admin"})
    async def microsoft_admin(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_service_health', 'get_service_health', 'list_service_health_issues', 'get_service_health_issue', 'list_service_update_messages', 'get_service_update_message', 'get_admin_sharepoint', 'update_admin_sharepoint', 'list_delegated_admin_relationships', 'get_delegated_admin_relationship'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft admin operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_service_health":
            return client.list_service_health(**kwargs)
        if action == "get_service_health":
            return client.get_service_health(**kwargs)
        if action == "list_service_health_issues":
            return client.list_service_health_issues(**kwargs)
        if action == "get_service_health_issue":
            return client.get_service_health_issue(**kwargs)
        if action == "list_service_update_messages":
            return client.list_service_update_messages(**kwargs)
        if action == "get_service_update_message":
            return client.get_service_update_message(**kwargs)
        if action == "get_admin_sharepoint":
            return client.get_admin_sharepoint(**kwargs)
        if action == "update_admin_sharepoint":
            return client.update_admin_sharepoint(**kwargs)
        if action == "list_delegated_admin_relationships":
            return client.list_delegated_admin_relationships(**kwargs)
        if action == "get_delegated_admin_relationship":
            return client.get_delegated_admin_relationship(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_organization_tools(mcp: FastMCP):
    @mcp.tool(tags={"organization"})
    async def microsoft_organization(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_organization', 'get_organization', 'update_organization', 'get_org_branding', 'update_org_branding'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft organization operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_organization":
            return client.list_organization(**kwargs)
        if action == "get_organization":
            return client.get_organization(**kwargs)
        if action == "update_organization":
            return client.update_organization(**kwargs)
        if action == "get_org_branding":
            return client.get_org_branding(**kwargs)
        if action == "update_org_branding":
            return client.update_org_branding(**kwargs)
        raise ValueError(f"Unknown action: {action}")


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


def register_subscriptions_tools(mcp: FastMCP):
    @mcp.tool(tags={"subscriptions"})
    async def microsoft_subscriptions(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_subscriptions', 'get_subscription', 'create_subscription', 'update_subscription', 'delete_subscription'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft subscriptions operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_subscriptions":
            return client.list_subscriptions(**kwargs)
        if action == "get_subscription":
            return client.get_subscription(**kwargs)
        if action == "create_subscription":
            return client.create_subscription(**kwargs)
        if action == "update_subscription":
            return client.update_subscription(**kwargs)
        if action == "delete_subscription":
            return client.delete_subscription(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_communications_tools(mcp: FastMCP):
    @mcp.tool(tags={"communications"})
    async def microsoft_communications(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_online_meetings', 'get_online_meeting', 'create_online_meeting', 'update_online_meeting', 'delete_online_meeting', 'list_call_records', 'get_call_record', 'list_presences', 'get_presence', 'get_my_presence'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft communications operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_online_meetings":
            return client.list_online_meetings(**kwargs)
        if action == "get_online_meeting":
            return client.get_online_meeting(**kwargs)
        if action == "create_online_meeting":
            return client.create_online_meeting(**kwargs)
        if action == "update_online_meeting":
            return client.update_online_meeting(**kwargs)
        if action == "delete_online_meeting":
            return client.delete_online_meeting(**kwargs)
        if action == "list_call_records":
            return client.list_call_records(**kwargs)
        if action == "get_call_record":
            return client.get_call_record(**kwargs)
        if action == "list_presences":
            return client.list_presences(**kwargs)
        if action == "get_presence":
            return client.get_presence(**kwargs)
        if action == "get_my_presence":
            return client.get_my_presence(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_identity_tools(mcp: FastMCP):
    @mcp.tool(tags={"identity"})
    async def microsoft_identity(
        action: str = Field(
            description="Action to perform. Must be one of: 'create_invitation', 'list_conditional_access_policies', 'get_conditional_access_policy', 'create_conditional_access_policy', 'update_conditional_access_policy', 'delete_conditional_access_policy', 'list_access_reviews', 'get_access_review', 'list_entitlement_access_packages', 'list_lifecycle_workflows'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft identity operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "create_invitation":
            return client.create_invitation(**kwargs)
        if action == "list_conditional_access_policies":
            return client.list_conditional_access_policies(**kwargs)
        if action == "get_conditional_access_policy":
            return client.get_conditional_access_policy(**kwargs)
        if action == "create_conditional_access_policy":
            return client.create_conditional_access_policy(**kwargs)
        if action == "update_conditional_access_policy":
            return client.update_conditional_access_policy(**kwargs)
        if action == "delete_conditional_access_policy":
            return client.delete_conditional_access_policy(**kwargs)
        if action == "list_access_reviews":
            return client.list_access_reviews(**kwargs)
        if action == "get_access_review":
            return client.get_access_review(**kwargs)
        if action == "list_entitlement_access_packages":
            return client.list_entitlement_access_packages(**kwargs)
        if action == "list_lifecycle_workflows":
            return client.list_lifecycle_workflows(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_security_tools(mcp: FastMCP):
    @mcp.tool(tags={"security"})
    async def microsoft_security(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_security_alerts', 'get_security_alert', 'update_security_alert', 'list_security_incidents', 'get_security_incident', 'update_security_incident', 'list_secure_scores', 'list_threat_intelligence_hosts', 'get_threat_intelligence_host', 'run_hunting_query', 'list_risk_detections', 'get_risk_detection', 'list_risky_users', 'get_risky_user', 'dismiss_risky_user', 'list_sensitivity_labels', 'get_sensitivity_label'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft security operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_security_alerts":
            return client.list_security_alerts(**kwargs)
        if action == "get_security_alert":
            return client.get_security_alert(**kwargs)
        if action == "update_security_alert":
            return client.update_security_alert(**kwargs)
        if action == "list_security_incidents":
            return client.list_security_incidents(**kwargs)
        if action == "get_security_incident":
            return client.get_security_incident(**kwargs)
        if action == "update_security_incident":
            return client.update_security_incident(**kwargs)
        if action == "list_secure_scores":
            return client.list_secure_scores(**kwargs)
        if action == "list_threat_intelligence_hosts":
            return client.list_threat_intelligence_hosts(**kwargs)
        if action == "get_threat_intelligence_host":
            return client.get_threat_intelligence_host(**kwargs)
        if action == "run_hunting_query":
            return client.run_hunting_query(**kwargs)
        if action == "list_risk_detections":
            return client.list_risk_detections(**kwargs)
        if action == "get_risk_detection":
            return client.get_risk_detection(**kwargs)
        if action == "list_risky_users":
            return client.list_risky_users(**kwargs)
        if action == "get_risky_user":
            return client.get_risky_user(**kwargs)
        if action == "dismiss_risky_user":
            return client.dismiss_risky_user(**kwargs)
        if action == "list_sensitivity_labels":
            return client.list_sensitivity_labels(**kwargs)
        if action == "get_sensitivity_label":
            return client.get_sensitivity_label(**kwargs)
        raise ValueError(f"Unknown action: {action}")


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


def register_reports_tools(mcp: FastMCP):
    @mcp.tool(tags={"reports"})
    async def microsoft_reports(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_email_activity_report', 'get_mailbox_usage_report', 'get_office365_active_users', 'get_sharepoint_activity_report', 'get_teams_user_activity', 'get_onedrive_usage_report'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft reports operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_email_activity_report":
            return client.get_email_activity_report(**kwargs)
        if action == "get_mailbox_usage_report":
            return client.get_mailbox_usage_report(**kwargs)
        if action == "get_office365_active_users":
            return client.get_office365_active_users(**kwargs)
        if action == "get_sharepoint_activity_report":
            return client.get_sharepoint_activity_report(**kwargs)
        if action == "get_teams_user_activity":
            return client.get_teams_user_activity(**kwargs)
        if action == "get_onedrive_usage_report":
            return client.get_onedrive_usage_report(**kwargs)
        raise ValueError(f"Unknown action: {action}")


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
            return client.list_applications(**kwargs)
        if action == "get_application":
            return client.get_application(**kwargs)
        if action == "create_application":
            return client.create_application(**kwargs)
        if action == "update_application":
            return client.update_application(**kwargs)
        if action == "delete_application":
            return client.delete_application(**kwargs)
        if action == "add_application_password":
            return client.add_application_password(**kwargs)
        if action == "remove_application_password":
            return client.remove_application_password(**kwargs)
        if action == "list_service_principals":
            return client.list_service_principals(**kwargs)
        if action == "get_service_principal":
            return client.get_service_principal(**kwargs)
        if action == "create_service_principal":
            return client.create_service_principal(**kwargs)
        if action == "update_service_principal":
            return client.update_service_principal(**kwargs)
        if action == "delete_service_principal":
            return client.delete_service_principal(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_directory_tools(mcp: FastMCP):
    @mcp.tool(tags={"directory"})
    async def microsoft_directory(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_directory_objects', 'get_directory_object', 'list_directory_roles', 'get_directory_role', 'list_directory_role_templates', 'list_deleted_items', 'restore_deleted_item', 'list_role_definitions', 'get_role_definition', 'list_role_assignments', 'get_role_assignment', 'create_role_assignment'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft directory operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_directory_objects":
            return client.list_directory_objects(**kwargs)
        if action == "get_directory_object":
            return client.get_directory_object(**kwargs)
        if action == "list_directory_roles":
            return client.list_directory_roles(**kwargs)
        if action == "get_directory_role":
            return client.get_directory_role(**kwargs)
        if action == "list_directory_role_templates":
            return client.list_directory_role_templates(**kwargs)
        if action == "list_deleted_items":
            return client.list_deleted_items(**kwargs)
        if action == "restore_deleted_item":
            return client.restore_deleted_item(**kwargs)
        if action == "list_role_definitions":
            return client.list_role_definitions(**kwargs)
        if action == "get_role_definition":
            return client.get_role_definition(**kwargs)
        if action == "list_role_assignments":
            return client.list_role_assignments(**kwargs)
        if action == "get_role_assignment":
            return client.get_role_assignment(**kwargs)
        if action == "create_role_assignment":
            return client.create_role_assignment(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_policies_tools(mcp: FastMCP):
    @mcp.tool(tags={"policies"})
    async def microsoft_policies(
        action: str = Field(
            description="Action to perform. Must be one of: 'get_authorization_policy', 'list_token_lifetime_policies', 'list_token_issuance_policies', 'list_permission_grant_policies', 'get_admin_consent_policy'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft policies operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "get_authorization_policy":
            return client.get_authorization_policy(**kwargs)
        if action == "list_token_lifetime_policies":
            return client.list_token_lifetime_policies(**kwargs)
        if action == "list_token_issuance_policies":
            return client.list_token_issuance_policies(**kwargs)
        if action == "list_permission_grant_policies":
            return client.list_permission_grant_policies(**kwargs)
        if action == "get_admin_consent_policy":
            return client.get_admin_consent_policy(**kwargs)
        raise ValueError(f"Unknown action: {action}")


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


def register_education_tools(mcp: FastMCP):
    @mcp.tool(tags={"education"})
    async def microsoft_education(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_education_classes', 'get_education_class', 'list_education_schools', 'get_education_school', 'list_education_users', 'list_education_assignments'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft education operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_education_classes":
            return client.list_education_classes(**kwargs)
        if action == "get_education_class":
            return client.get_education_class(**kwargs)
        if action == "list_education_schools":
            return client.list_education_schools(**kwargs)
        if action == "get_education_school":
            return client.get_education_school(**kwargs)
        if action == "list_education_users":
            return client.list_education_users(**kwargs)
        if action == "list_education_assignments":
            return client.list_education_assignments(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_agreements_tools(mcp: FastMCP):
    @mcp.tool(tags={"agreements"})
    async def microsoft_agreements(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_agreements', 'get_agreement', 'create_agreement', 'delete_agreement'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft agreements operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_agreements":
            return client.list_agreements(**kwargs)
        if action == "get_agreement":
            return client.get_agreement(**kwargs)
        if action == "create_agreement":
            return client.create_agreement(**kwargs)
        if action == "delete_agreement":
            return client.delete_agreement(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_places_tools(mcp: FastMCP):
    @mcp.tool(tags={"places"})
    async def microsoft_places(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_rooms', 'list_room_lists', 'get_place', 'update_place'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft places operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_rooms":
            return client.list_rooms(**kwargs)
        if action == "list_room_lists":
            return client.list_room_lists(**kwargs)
        if action == "get_place":
            return client.get_place(**kwargs)
        if action == "update_place":
            return client.update_place(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_print_tools(mcp: FastMCP):
    @mcp.tool(tags={"print"})
    async def microsoft_print(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_printers', 'get_printer', 'list_print_jobs', 'create_print_job', 'list_print_shares'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
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
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_printers":
            return client.list_printers(**kwargs)
        if action == "get_printer":
            return client.get_printer(**kwargs)
        if action == "list_print_jobs":
            return client.list_print_jobs(**kwargs)
        if action == "create_print_job":
            return client.create_print_job(**kwargs)
        if action == "list_print_shares":
            return client.list_print_shares(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_privacy_tools(mcp: FastMCP):
    @mcp.tool(tags={"privacy"})
    async def microsoft_privacy(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_subject_rights_requests', 'get_subject_rights_request', 'create_subject_rights_request'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft privacy operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_subject_rights_requests":
            return client.list_subject_rights_requests(**kwargs)
        if action == "get_subject_rights_request":
            return client.get_subject_rights_request(**kwargs)
        if action == "create_subject_rights_request":
            return client.create_subject_rights_request(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_solutions_tools(mcp: FastMCP):
    @mcp.tool(tags={"solutions"})
    async def microsoft_solutions(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_booking_businesses', 'get_booking_business', 'list_booking_appointments', 'create_booking_appointment', 'list_virtual_events'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft solutions operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_booking_businesses":
            return client.list_booking_businesses(**kwargs)
        if action == "get_booking_business":
            return client.get_booking_business(**kwargs)
        if action == "list_booking_appointments":
            return client.list_booking_appointments(**kwargs)
        if action == "create_booking_appointment":
            return client.create_booking_appointment(**kwargs)
        if action == "list_virtual_events":
            return client.list_virtual_events(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_storage_tools(mcp: FastMCP):
    @mcp.tool(tags={"storage"})
    async def microsoft_storage(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_file_storage_containers', 'get_file_storage_container', 'create_file_storage_container'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft storage operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_file_storage_containers":
            return client.list_file_storage_containers(**kwargs)
        if action == "get_file_storage_container":
            return client.get_file_storage_container(**kwargs)
        if action == "create_file_storage_container":
            return client.create_file_storage_container(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_employee_experience_tools(mcp: FastMCP):
    @mcp.tool(tags={"employee_experience"})
    async def microsoft_employee_experience(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_learning_providers', 'get_learning_provider', 'list_learning_course_activities'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft employee experience operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_learning_providers":
            return client.list_learning_providers(**kwargs)
        if action == "get_learning_provider":
            return client.get_learning_provider(**kwargs)
        if action == "list_learning_course_activities":
            return client.list_learning_course_activities(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def register_connections_tools(mcp: FastMCP):
    @mcp.tool(tags={"connections"})
    async def microsoft_connections(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_external_connections', 'get_external_connection', 'create_external_connection', 'delete_external_connection'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft connections operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_external_connections":
            return client.list_external_connections(**kwargs)
        if action == "get_external_connection":
            return client.get_external_connection(**kwargs)
        if action == "create_external_connection":
            return client.create_external_connection(**kwargs)
        if action == "delete_external_connection":
            return client.delete_external_connection(**kwargs)
        raise ValueError(f"Unknown action: {action}")


def get_mcp_instance() -> tuple[Any, ...]:
    """Initialize and return the MCP instance."""
    load_dotenv(find_dotenv())
    args, mcp, middlewares = create_mcp_server(
        name="microsoft-agent MCP",
        version=__version__,
        instructions="microsoft-agent MCP Server — Condensed Action-Routed Tools.",
    )

    @mcp.custom_route("/health", methods=["GET"])
    async def health_check(request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    DEFAULT_AUTHTOOL = to_boolean(os.getenv("AUTHTOOL", "True"))
    if DEFAULT_AUTHTOOL:
        register_auth_tools(mcp)
    DEFAULT_METATOOL = to_boolean(os.getenv("METATOOL", "True"))
    if DEFAULT_METATOOL:
        register_meta_tools(mcp)
    DEFAULT_MAILTOOL = to_boolean(os.getenv("MAILTOOL", "True"))
    if DEFAULT_MAILTOOL:
        register_mail_tools(mcp)
    DEFAULT_FILESTOOL = to_boolean(os.getenv("FILESTOOL", "True"))
    if DEFAULT_FILESTOOL:
        register_files_tools(mcp)
    DEFAULT_CALENDARTOOL = to_boolean(os.getenv("CALENDARTOOL", "True"))
    if DEFAULT_CALENDARTOOL:
        register_calendar_tools(mcp)
    DEFAULT_NOTESTOOL = to_boolean(os.getenv("NOTESTOOL", "True"))
    if DEFAULT_NOTESTOOL:
        register_notes_tools(mcp)
    DEFAULT_TASKSTOOL = to_boolean(os.getenv("TASKSTOOL", "True"))
    if DEFAULT_TASKSTOOL:
        register_tasks_tools(mcp)
    DEFAULT_CONTACTSTOOL = to_boolean(os.getenv("CONTACTSTOOL", "True"))
    if DEFAULT_CONTACTSTOOL:
        register_contacts_tools(mcp)
    DEFAULT_USERTOOL = to_boolean(os.getenv("USERTOOL", "True"))
    if DEFAULT_USERTOOL:
        register_user_tools(mcp)
    DEFAULT_CHATTOOL = to_boolean(os.getenv("CHATTOOL", "True"))
    if DEFAULT_CHATTOOL:
        register_chat_tools(mcp)
    DEFAULT_TEAMSTOOL = to_boolean(os.getenv("TEAMSTOOL", "True"))
    if DEFAULT_TEAMSTOOL:
        register_teams_tools(mcp)
    DEFAULT_SITESTOOL = to_boolean(os.getenv("SITESTOOL", "True"))
    if DEFAULT_SITESTOOL:
        register_sites_tools(mcp)
    DEFAULT_SEARCHTOOL = to_boolean(os.getenv("SEARCHTOOL", "True"))
    if DEFAULT_SEARCHTOOL:
        register_search_tools(mcp)
    DEFAULT_GROUPSTOOL = to_boolean(os.getenv("GROUPSTOOL", "True"))
    if DEFAULT_GROUPSTOOL:
        register_groups_tools(mcp)
    DEFAULT_ADMINTOOL = to_boolean(os.getenv("ADMINTOOL", "True"))
    if DEFAULT_ADMINTOOL:
        register_admin_tools(mcp)
    DEFAULT_ORGANIZATIONTOOL = to_boolean(os.getenv("ORGANIZATIONTOOL", "True"))
    if DEFAULT_ORGANIZATIONTOOL:
        register_organization_tools(mcp)
    DEFAULT_DOMAINSTOOL = to_boolean(os.getenv("DOMAINSTOOL", "True"))
    if DEFAULT_DOMAINSTOOL:
        register_domains_tools(mcp)
    DEFAULT_SUBSCRIPTIONSTOOL = to_boolean(os.getenv("SUBSCRIPTIONSTOOL", "True"))
    if DEFAULT_SUBSCRIPTIONSTOOL:
        register_subscriptions_tools(mcp)
    DEFAULT_COMMUNICATIONSTOOL = to_boolean(os.getenv("COMMUNICATIONSTOOL", "True"))
    if DEFAULT_COMMUNICATIONSTOOL:
        register_communications_tools(mcp)
    DEFAULT_IDENTITYTOOL = to_boolean(os.getenv("IDENTITYTOOL", "True"))
    if DEFAULT_IDENTITYTOOL:
        register_identity_tools(mcp)
    DEFAULT_SECURITYTOOL = to_boolean(os.getenv("SECURITYTOOL", "True"))
    if DEFAULT_SECURITYTOOL:
        register_security_tools(mcp)
    DEFAULT_AUDITTOOL = to_boolean(os.getenv("AUDITTOOL", "True"))
    if DEFAULT_AUDITTOOL:
        register_audit_tools(mcp)
    DEFAULT_REPORTSTOOL = to_boolean(os.getenv("REPORTSTOOL", "True"))
    if DEFAULT_REPORTSTOOL:
        register_reports_tools(mcp)
    DEFAULT_APPLICATIONSTOOL = to_boolean(os.getenv("APPLICATIONSTOOL", "True"))
    if DEFAULT_APPLICATIONSTOOL:
        register_applications_tools(mcp)
    DEFAULT_DIRECTORYTOOL = to_boolean(os.getenv("DIRECTORYTOOL", "True"))
    if DEFAULT_DIRECTORYTOOL:
        register_directory_tools(mcp)
    DEFAULT_POLICIESTOOL = to_boolean(os.getenv("POLICIESTOOL", "True"))
    if DEFAULT_POLICIESTOOL:
        register_policies_tools(mcp)
    DEFAULT_DEVICESTOOL = to_boolean(os.getenv("DEVICESTOOL", "True"))
    if DEFAULT_DEVICESTOOL:
        register_devices_tools(mcp)
    DEFAULT_EDUCATIONTOOL = to_boolean(os.getenv("EDUCATIONTOOL", "True"))
    if DEFAULT_EDUCATIONTOOL:
        register_education_tools(mcp)
    DEFAULT_AGREEMENTSTOOL = to_boolean(os.getenv("AGREEMENTSTOOL", "True"))
    if DEFAULT_AGREEMENTSTOOL:
        register_agreements_tools(mcp)
    DEFAULT_PLACESTOOL = to_boolean(os.getenv("PLACESTOOL", "True"))
    if DEFAULT_PLACESTOOL:
        register_places_tools(mcp)
    DEFAULT_PRINTTOOL = to_boolean(os.getenv("PRINTTOOL", "True"))
    if DEFAULT_PRINTTOOL:
        register_print_tools(mcp)
    DEFAULT_PRIVACYTOOL = to_boolean(os.getenv("PRIVACYTOOL", "True"))
    if DEFAULT_PRIVACYTOOL:
        register_privacy_tools(mcp)
    DEFAULT_SOLUTIONSTOOL = to_boolean(os.getenv("SOLUTIONSTOOL", "True"))
    if DEFAULT_SOLUTIONSTOOL:
        register_solutions_tools(mcp)
    DEFAULT_STORAGETOOL = to_boolean(os.getenv("STORAGETOOL", "True"))
    if DEFAULT_STORAGETOOL:
        register_storage_tools(mcp)
    DEFAULT_EMPLOYEE_EXPERIENCETOOL = to_boolean(
        os.getenv("EMPLOYEE_EXPERIENCETOOL", "True")
    )
    if DEFAULT_EMPLOYEE_EXPERIENCETOOL:
        register_employee_experience_tools(mcp)
    DEFAULT_CONNECTIONSTOOL = to_boolean(os.getenv("CONNECTIONSTOOL", "True"))
    if DEFAULT_CONNECTIONSTOOL:
        register_connections_tools(mcp)

    for mw in middlewares:
        mcp.add_middleware(mw)
    return mcp, args, middlewares


def mcp_server() -> None:
    mcp, args, middlewares = get_mcp_instance()
    print(f"microsoft-agent MCP v{__version__}", file=sys.stderr)
    print("\nStarting MCP Server", file=sys.stderr)
    print(f"  Transport: {args.transport.upper()}", file=sys.stderr)
    print(f"  Auth: {args.auth_type}", file=sys.stderr)

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "streamable-http":
        mcp.run(transport="streamable-http", host=args.host, port=args.port)
    elif args.transport == "sse":
        mcp.run(transport="sse", host=args.host, port=args.port)
    else:
        logger.error("Invalid transport", extra={"transport": args.transport})
        sys.exit(1)


if __name__ == "__main__":
    mcp_server()
