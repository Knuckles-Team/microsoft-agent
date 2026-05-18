#!/usr/bin/python
import warnings

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
from fastmcp import FastMCP
from fastmcp.dependencies import Depends
from fastmcp.utilities.logging import get_logger
from pydantic import Field
from starlette.requests import Request
from starlette.responses import JSONResponse

from microsoft_agent.auth import get_client

__version__ = "0.11.0"

logger = get_logger(name="microsoft-agent")
logger.setLevel(logging.INFO)


def register_auth_tools(mcp: FastMCP):
    @mcp.tool(tags={"auth"})
    async def microsoft_auth(
        action: str = Field(description="Action to perform. Must be one of: 'login', 'logout', 'verify_login', 'list_accounts'"),
        force: bool | None = Field(default=None, description="force"),
        client=Depends(get_client),
    ) -> dict:
        """Manage auth operations.

        Actions:
          - 'login': Authenticate with Microsoft.
          - 'logout': Logout.
          - 'verify_login': Verify login status.
          - 'list_accounts': List accounts.
        """
        kwargs: dict[str, Any]
        if action == "login":
            kwargs = {"force": force}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.login(**kwargs)
        if action == "logout":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.logout(**kwargs)
        if action == "verify_login":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.verify_login(**kwargs)
        if action == "list_accounts":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_accounts(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: login', 'logout', 'verify_login', 'list_accounts")


def register_meta_tools(mcp: FastMCP):
    @mcp.tool(tags={"meta"})
    async def microsoft_meta(
        action: str = Field(description="Action to perform. Must be one of: 'searches'"),
        client=Depends(get_client),
    ) -> dict:
        """Manage meta operations.

        Actions:
          - 'searches': Call searches
        """
        kwargs: dict[str, Any]
        if action == "searches":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.searches(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: searches")


def register_mail_tools(mcp: FastMCP):
    @mcp.tool(tags={"mail"})
    async def microsoft_mail(
        action: str = Field(description="Action to perform. Must be one of: 'list_mail_messages', 'list_mail_folders', 'list_mail_folder_messages', 'get_mail_message', 'send_mail', 'list_shared_mailbox_messages', 'list_shared_mailbox_folder_messages', 'get_shared_mailbox_message', 'send_shared_mailbox_mail', 'create_draft_email', 'delete_mail_message', 'move_mail_message', 'update_mail_message', 'add_mail_attachment', 'list_mail_attachments', 'get_mail_attachment', 'delete_mail_attachment', 'get_root_folder', 'list_folder_files', 'list_chat_messages', 'get_chat_message', 'send_chat_message', 'list_channel_messages', 'get_channel_message', 'send_channel_message', 'list_chat_message_replies', 'reply_to_chat_message'"),
        params: dict | None = Field(default=None, description="params"),
        mailFolder_id: str | None = Field(default=None, description="mailFolder id"),
        message_id: str | None = Field(default=None, description="message id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        attachment_id: str | None = Field(default=None, description="attachment id"),
        user_id: str | None = Field(default=None, description="user id"),
        drive_id: str | None = Field(default=None, description="drive id"),
        driveItem_id: str | None = Field(default=None, description="driveItem id"),
        chat_id: str | None = Field(default=None, description="chat id"),
        chatMessage_id: str | None = Field(default=None, description="chatMessage id"),
        team_id: str | None = Field(default=None, description="team id"),
        channel_id: str | None = Field(default=None, description="channel id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage mail operations.

        Actions:
          - 'list_mail_messages': List mail messages.
          - 'list_mail_folders': List mail folders.
          - 'list_mail_folder_messages': List messages in a specific folder.
          - 'get_mail_message': Get a specific message.
          - 'send_mail': Send mail.
          - 'list_shared_mailbox_messages': List messages in a shared mailbox.
          - 'list_shared_mailbox_folder_messages': List messages in a shared mailbox folder.
          - 'get_shared_mailbox_message': Get a message from a shared mailbox.
          - 'send_shared_mailbox_mail': Send mail from a shared mailbox.
          - 'create_draft_email': Create draft email.
          - 'delete_mail_message': Delete a message.
          - 'move_mail_message': Move a message to a folder.
          - 'update_mail_message': Update a message.
          - 'add_mail_attachment': Add attachment to message.
          - 'list_mail_attachments': List attachments.
          - 'get_mail_attachment': Get attachment.
          - 'delete_mail_attachment': Delete attachment.
          - 'get_root_folder': Alias for get_drive_root_item.
          - 'list_folder_files': List folder files.
          - 'list_chat_messages': List chat messages.
          - 'get_chat_message': Get chat message.
          - 'send_chat_message': Send chat message.
          - 'list_channel_messages': List channel messages.
          - 'get_channel_message': Get channel message.
          - 'send_channel_message': Send channel message.
          - 'list_chat_message_replies': List chat message replies.
          - 'reply_to_chat_message': Reply to a chat message.
        """
        kwargs: dict[str, Any]
        if action == "list_mail_messages":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_mail_messages(**kwargs)
        if action == "list_mail_folders":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_mail_folders(**kwargs)
        if action == "list_mail_folder_messages":
            kwargs = {"mailFolder_id": mailFolder_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_mail_folder_messages(**kwargs)
        if action == "get_mail_message":
            kwargs = {"message_id": message_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_mail_message(**kwargs)
        if action == "send_mail":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.send_mail(**kwargs)
        if action == "list_shared_mailbox_messages":
            kwargs = {"user_id": user_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_shared_mailbox_messages(**kwargs)
        if action == "list_shared_mailbox_folder_messages":
            kwargs = {"user_id": user_id, "mailFolder_id": mailFolder_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_shared_mailbox_folder_messages(**kwargs)
        if action == "get_shared_mailbox_message":
            kwargs = {"user_id": user_id, "message_id": message_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_shared_mailbox_message(**kwargs)
        if action == "send_shared_mailbox_mail":
            kwargs = {"user_id": user_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.send_shared_mailbox_mail(**kwargs)
        if action == "create_draft_email":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_draft_email(**kwargs)
        if action == "delete_mail_message":
            kwargs = {"message_id": message_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_mail_message(**kwargs)
        if action == "move_mail_message":
            kwargs = {"message_id": message_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.move_mail_message(**kwargs)
        if action == "update_mail_message":
            kwargs = {"message_id": message_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_mail_message(**kwargs)
        if action == "add_mail_attachment":
            kwargs = {"message_id": message_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.add_mail_attachment(**kwargs)
        if action == "list_mail_attachments":
            kwargs = {"message_id": message_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_mail_attachments(**kwargs)
        if action == "get_mail_attachment":
            kwargs = {"message_id": message_id, "attachment_id": attachment_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_mail_attachment(**kwargs)
        if action == "delete_mail_attachment":
            kwargs = {"message_id": message_id, "attachment_id": attachment_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_mail_attachment(**kwargs)
        if action == "get_root_folder":
            kwargs = {"drive_id": drive_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_root_folder(**kwargs)
        if action == "list_folder_files":
            kwargs = {"drive_id": drive_id, "driveItem_id": driveItem_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_folder_files(**kwargs)
        if action == "list_chat_messages":
            kwargs = {"chat_id": chat_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_chat_messages(**kwargs)
        if action == "get_chat_message":
            kwargs = {"chat_id": chat_id, "chatMessage_id": chatMessage_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_chat_message(**kwargs)
        if action == "send_chat_message":
            kwargs = {"chat_id": chat_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.send_chat_message(**kwargs)
        if action == "list_channel_messages":
            kwargs = {"team_id": team_id, "channel_id": channel_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_channel_messages(**kwargs)
        if action == "get_channel_message":
            kwargs = {"team_id": team_id, "channel_id": channel_id, "chatMessage_id": chatMessage_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_channel_message(**kwargs)
        if action == "send_channel_message":
            kwargs = {"team_id": team_id, "channel_id": channel_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.send_channel_message(**kwargs)
        if action == "list_chat_message_replies":
            kwargs = {"chat_id": chat_id, "chatMessage_id": chatMessage_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_chat_message_replies(**kwargs)
        if action == "reply_to_chat_message":
            kwargs = {"chat_id": chat_id, "chatMessage_id": chatMessage_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.reply_to_chat_message(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_mail_messages', 'list_mail_folders', 'list_mail_folder_messages', 'get_mail_message', 'send_mail', 'list_shared_mailbox_messages', 'list_shared_mailbox_folder_messages', 'get_shared_mailbox_message', 'send_shared_mailbox_mail', 'create_draft_email', 'delete_mail_message', 'move_mail_message', 'update_mail_message', 'add_mail_attachment', 'list_mail_attachments', 'get_mail_attachment', 'delete_mail_attachment', 'get_root_folder', 'list_folder_files', 'list_chat_messages', 'get_chat_message', 'send_chat_message', 'list_channel_messages', 'get_channel_message', 'send_channel_message', 'list_chat_message_replies', 'reply_to_chat_message")


def register_files_tools(mcp: FastMCP):
    @mcp.tool(tags={"files"})
    async def microsoft_files(
        action: str = Field(description="Action to perform. Must be one of: 'list_users', 'list_drives', 'get_drive_root_item', 'download_onedrive_file_content', 'delete_onedrive_file', 'upload_file_content', 'create_excel_chart', 'format_excel_range', 'sort_excel_range', 'get_excel_range', 'list_excel_worksheets', 'list_excel_tables', 'get_excel_workbook', 'list_onenote_notebooks', 'list_onenote_notebook_sections', 'list_onenote_section_pages', 'list_todo_task_lists', 'list_todo_tasks', 'list_planner_tasks', 'list_plan_tasks', 'list_outlook_contacts', 'list_chats', 'get_excel_worksheet', 'list_joined_teams', 'list_team_channels', 'list_team_members', 'list_site_drives', 'get_site_drive_by_id', 'list_site_items', 'get_site_item', 'list_site_lists', 'get_site_list', 'list_sharepoint_site_list_items', 'get_sharepoint_site_list_item', 'get_excel_table'"),
        params: dict | None = Field(default=None, description="params"),
        drive_id: str | None = Field(default=None, description="drive id"),
        driveItem_id: str | None = Field(default=None, description="driveItem id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        site_id: str | None = Field(default=None, description="site id"),
        list_id: str | None = Field(default=None, description="list id"),
        item_id: str | None = Field(default=None, description="item id"),
        worksheet_id: str | None = Field(default=None, description="worksheet id"),
        table_id: str | None = Field(default=None, description="table id"),
        notebook_id: str | None = Field(default=None, description="notebook id"),
        onenoteSection_id: str | None = Field(default=None, description="onenoteSection id"),
        todoTaskList_id: str | None = Field(default=None, description="todoTaskList id"),
        plannerPlan_id: str | None = Field(default=None, description="plannerPlan id"),
        team_id: str | None = Field(default=None, description="team id"),
        listItem_id: str | None = Field(default=None, description="listItem id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage files operations.

        Actions:
          - 'list_users': List users.
          - 'list_drives': List drives.
          - 'get_drive_root_item': Get drive root item.
          - 'download_onedrive_file_content': Download file content.
          - 'delete_onedrive_file': Delete file.
          - 'upload_file_content': Upload file content.
          - 'create_excel_chart': Call create_excel_chart
          - 'format_excel_range': Call format_excel_range
          - 'sort_excel_range': Call sort_excel_range
          - 'get_excel_range': Call get_excel_range
          - 'list_excel_worksheets': List Excel worksheets.
          - 'list_excel_tables': List Excel tables.
          - 'get_excel_workbook': Get Excel workbook.
          - 'list_onenote_notebooks': Call list_onenote_notebooks
          - 'list_onenote_notebook_sections': List Onenote notebook sections.
          - 'list_onenote_section_pages': List Onenote section pages.
          - 'list_todo_task_lists': List Todo task lists.
          - 'list_todo_tasks': List Todo tasks.
          - 'list_planner_tasks': List Planner tasks.
          - 'list_plan_tasks': List tasks for a Planner plan.
          - 'list_outlook_contacts': List Outlook contacts.
          - 'list_chats': List user chats.
          - 'get_excel_worksheet': Get Excel worksheet.
          - 'list_joined_teams': List joined teams.
          - 'list_team_channels': List team channels.
          - 'list_team_members': List team members.
          - 'list_site_drives': List drives for a SharePoint site.
          - 'get_site_drive_by_id': Call get_site_drive_by_id
          - 'list_site_items': Call list_site_items
          - 'get_site_item': Call get_site_item
          - 'list_site_lists': List lists for a SharePoint site.
          - 'get_site_list': Get a SharePoint site list.
          - 'list_sharepoint_site_list_items': List items in a SharePoint site list.
          - 'get_sharepoint_site_list_item': Get an item in a SharePoint site list.
          - 'get_excel_table': Get Excel table.
        """
        kwargs: dict[str, Any]
        if action == "list_users":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_users(**kwargs)
        if action == "list_drives":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_drives(**kwargs)
        if action == "get_drive_root_item":
            kwargs = {"drive_id": drive_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_drive_root_item(**kwargs)
        if action == "download_onedrive_file_content":
            kwargs = {"drive_id": drive_id, "driveItem_id": driveItem_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.download_onedrive_file_content(**kwargs)
        if action == "delete_onedrive_file":
            kwargs = {"drive_id": drive_id, "driveItem_id": driveItem_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_onedrive_file(**kwargs)
        if action == "upload_file_content":
            kwargs = {"drive_id": drive_id, "driveItem_id": driveItem_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.upload_file_content(**kwargs)
        if action == "create_excel_chart":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_excel_chart(**kwargs)
        if action == "format_excel_range":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.format_excel_range(**kwargs)
        if action == "sort_excel_range":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.sort_excel_range(**kwargs)
        if action == "get_excel_range":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_excel_range(**kwargs)
        if action == "list_excel_worksheets":
            kwargs = {"drive_id": drive_id, "item_id": item_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_excel_worksheets(**kwargs)
        if action == "list_excel_tables":
            kwargs = {"drive_id": drive_id, "item_id": item_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_excel_tables(**kwargs)
        if action == "get_excel_workbook":
            kwargs = {"drive_id": drive_id, "item_id": item_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_excel_workbook(**kwargs)
        if action == "list_onenote_notebooks":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_onenote_notebooks(**kwargs)
        if action == "list_onenote_notebook_sections":
            kwargs = {"notebook_id": notebook_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_onenote_notebook_sections(**kwargs)
        if action == "list_onenote_section_pages":
            kwargs = {"onenoteSection_id": onenoteSection_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_onenote_section_pages(**kwargs)
        if action == "list_todo_task_lists":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_todo_task_lists(**kwargs)
        if action == "list_todo_tasks":
            kwargs = {"todoTaskList_id": todoTaskList_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_todo_tasks(**kwargs)
        if action == "list_planner_tasks":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_planner_tasks(**kwargs)
        if action == "list_plan_tasks":
            kwargs = {"plannerPlan_id": plannerPlan_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_plan_tasks(**kwargs)
        if action == "list_outlook_contacts":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_outlook_contacts(**kwargs)
        if action == "list_chats":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_chats(**kwargs)
        if action == "get_excel_worksheet":
            kwargs = {"drive_id": drive_id, "item_id": item_id, "worksheet_id": worksheet_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_excel_worksheet(**kwargs)
        if action == "list_joined_teams":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_joined_teams(**kwargs)
        if action == "list_team_channels":
            kwargs = {"team_id": team_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_team_channels(**kwargs)
        if action == "list_team_members":
            kwargs = {"team_id": team_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_team_members(**kwargs)
        if action == "list_site_drives":
            kwargs = {"site_id": site_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_site_drives(**kwargs)
        if action == "get_site_drive_by_id":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_site_drive_by_id(**kwargs)
        if action == "list_site_items":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_site_items(**kwargs)
        if action == "get_site_item":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_site_item(**kwargs)
        if action == "list_site_lists":
            kwargs = {"site_id": site_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_site_lists(**kwargs)
        if action == "get_site_list":
            kwargs = {"site_id": site_id, "list_id": list_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_site_list(**kwargs)
        if action == "list_sharepoint_site_list_items":
            kwargs = {"site_id": site_id, "list_id": list_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_sharepoint_site_list_items(**kwargs)
        if action == "get_sharepoint_site_list_item":
            kwargs = {"site_id": site_id, "list_id": list_id, "listItem_id": listItem_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sharepoint_site_list_item(**kwargs)
        if action == "get_excel_table":
            kwargs = {"drive_id": drive_id, "item_id": item_id, "table_id": table_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_excel_table(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_users', 'list_drives', 'get_drive_root_item', 'download_onedrive_file_content', 'delete_onedrive_file', 'upload_file_content', 'create_excel_chart', 'format_excel_range', 'sort_excel_range', 'get_excel_range', 'list_excel_worksheets', 'list_excel_tables', 'get_excel_workbook', 'list_onenote_notebooks', 'list_onenote_notebook_sections', 'list_onenote_section_pages', 'list_todo_task_lists', 'list_todo_tasks', 'list_planner_tasks', 'list_plan_tasks', 'list_outlook_contacts', 'list_chats', 'get_excel_worksheet', 'list_joined_teams', 'list_team_channels', 'list_team_members', 'list_site_drives', 'get_site_drive_by_id', 'list_site_items', 'get_site_item', 'list_site_lists', 'get_site_list', 'list_sharepoint_site_list_items', 'get_sharepoint_site_list_item', 'get_excel_table")


def register_calendar_tools(mcp: FastMCP):
    @mcp.tool(tags={"calendar"})
    async def microsoft_calendar(
        action: str = Field(description="Action to perform. Must be one of: 'list_calendar_events', 'get_calendar_event', 'create_calendar_event', 'update_calendar_event', 'delete_calendar_event', 'list_specific_calendar_events', 'get_specific_calendar_event', 'create_specific_calendar_event', 'update_specific_calendar_event', 'delete_specific_calendar_event', 'get_calendar_view', 'list_calendars', 'find_meeting_times'"),
        params: dict | None = Field(default=None, description="params"),
        timezone: str | None = Field(default=None, description="timezone"),
        event_id: str | None = Field(default=None, description="event id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        calendar_id: str | None = Field(default=None, description="calendar id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage calendar operations.

        Actions:
          - 'list_calendar_events': List calendar events.
          - 'get_calendar_event': Get calendar event.
          - 'create_calendar_event': Create calendar event.
          - 'update_calendar_event': Update calendar event.
          - 'delete_calendar_event': Delete calendar event.
          - 'list_specific_calendar_events': List events for a specific calendar.
          - 'get_specific_calendar_event': Get specific calendar event.
          - 'create_specific_calendar_event': Create specific calendar event.
          - 'update_specific_calendar_event': Update specific calendar event.
          - 'delete_specific_calendar_event': Delete specific calendar event.
          - 'get_calendar_view': Get calendar view.
          - 'list_calendars': List calendars.
          - 'find_meeting_times': Find meeting times.
        """
        kwargs: dict[str, Any]
        if action == "list_calendar_events":
            kwargs = {"params": params, "timezone": timezone}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_calendar_events(**kwargs)
        if action == "get_calendar_event":
            kwargs = {"event_id": event_id, "params": params, "timezone": timezone}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_calendar_event(**kwargs)
        if action == "create_calendar_event":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_calendar_event(**kwargs)
        if action == "update_calendar_event":
            kwargs = {"event_id": event_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_calendar_event(**kwargs)
        if action == "delete_calendar_event":
            kwargs = {"event_id": event_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_calendar_event(**kwargs)
        if action == "list_specific_calendar_events":
            kwargs = {"calendar_id": calendar_id, "params": params, "timezone": timezone}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_specific_calendar_events(**kwargs)
        if action == "get_specific_calendar_event":
            kwargs = {"calendar_id": calendar_id, "event_id": event_id, "params": params, "timezone": timezone}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_specific_calendar_event(**kwargs)
        if action == "create_specific_calendar_event":
            kwargs = {"calendar_id": calendar_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_specific_calendar_event(**kwargs)
        if action == "update_specific_calendar_event":
            kwargs = {"calendar_id": calendar_id, "event_id": event_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_specific_calendar_event(**kwargs)
        if action == "delete_specific_calendar_event":
            kwargs = {"calendar_id": calendar_id, "event_id": event_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_specific_calendar_event(**kwargs)
        if action == "get_calendar_view":
            kwargs = {"params": params, "timezone": timezone}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_calendar_view(**kwargs)
        if action == "list_calendars":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_calendars(**kwargs)
        if action == "find_meeting_times":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.find_meeting_times(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_calendar_events', 'get_calendar_event', 'create_calendar_event', 'update_calendar_event', 'delete_calendar_event', 'list_specific_calendar_events', 'get_specific_calendar_event', 'create_specific_calendar_event', 'update_specific_calendar_event', 'delete_specific_calendar_event', 'get_calendar_view', 'list_calendars', 'find_meeting_times")


def register_notes_tools(mcp: FastMCP):
    @mcp.tool(tags={"notes"})
    async def microsoft_notes(
        action: str = Field(description="Action to perform. Must be one of: 'get_onenote_page_content', 'create_onenote_page'"),
        onenotePage_id: str | None = Field(default=None, description="onenotePage id"),
        params: dict | None = Field(default=None, description="params"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage notes operations.

        Actions:
          - 'get_onenote_page_content': Get Onenote page content.
          - 'create_onenote_page': Create Onenote page.
        """
        kwargs: dict[str, Any]
        if action == "get_onenote_page_content":
            kwargs = {"onenotePage_id": onenotePage_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_onenote_page_content(**kwargs)
        if action == "create_onenote_page":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_onenote_page(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_onenote_page_content', 'create_onenote_page")


def register_tasks_tools(mcp: FastMCP):
    @mcp.tool(tags={"tasks"})
    async def microsoft_tasks(
        action: str = Field(description="Action to perform. Must be one of: 'get_todo_task', 'create_todo_task', 'update_todo_task', 'delete_todo_task', 'get_planner_plan', 'get_planner_task', 'create_planner_task', 'update_planner_task', 'update_planner_task_details'"),
        todoTaskList_id: str | None = Field(default=None, description="todoTaskList id"),
        todoTask_id: str | None = Field(default=None, description="todoTask id"),
        params: dict | None = Field(default=None, description="params"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        plannerPlan_id: str | None = Field(default=None, description="plannerPlan id"),
        plannerTask_id: str | None = Field(default=None, description="plannerTask id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage tasks operations.

        Actions:
          - 'get_todo_task': Get Todo task.
          - 'create_todo_task': Create Todo task.
          - 'update_todo_task': Update Todo task.
          - 'delete_todo_task': Delete Todo task.
          - 'get_planner_plan': Get Planner plan.
          - 'get_planner_task': Get Planner task.
          - 'create_planner_task': Create Planner task.
          - 'update_planner_task': Update Planner task.
          - 'update_planner_task_details': Update Planner task details.
        """
        kwargs: dict[str, Any]
        if action == "get_todo_task":
            kwargs = {"todoTaskList_id": todoTaskList_id, "todoTask_id": todoTask_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_todo_task(**kwargs)
        if action == "create_todo_task":
            kwargs = {"todoTaskList_id": todoTaskList_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_todo_task(**kwargs)
        if action == "update_todo_task":
            kwargs = {"todoTaskList_id": todoTaskList_id, "todoTask_id": todoTask_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_todo_task(**kwargs)
        if action == "delete_todo_task":
            kwargs = {"todoTaskList_id": todoTaskList_id, "todoTask_id": todoTask_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_todo_task(**kwargs)
        if action == "get_planner_plan":
            kwargs = {"plannerPlan_id": plannerPlan_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_planner_plan(**kwargs)
        if action == "get_planner_task":
            kwargs = {"plannerTask_id": plannerTask_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_planner_task(**kwargs)
        if action == "create_planner_task":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_planner_task(**kwargs)
        if action == "update_planner_task":
            kwargs = {"plannerTask_id": plannerTask_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_planner_task(**kwargs)
        if action == "update_planner_task_details":
            kwargs = {"plannerTask_id": plannerTask_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_planner_task_details(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_todo_task', 'create_todo_task', 'update_todo_task', 'delete_todo_task', 'get_planner_plan', 'get_planner_task', 'create_planner_task', 'update_planner_task', 'update_planner_task_details")


def register_contacts_tools(mcp: FastMCP):
    @mcp.tool(tags={"contacts"})
    async def microsoft_contacts(
        action: str = Field(description="Action to perform. Must be one of: 'get_outlook_contact', 'create_outlook_contact', 'update_outlook_contact', 'delete_outlook_contact'"),
        contact_id: str | None = Field(default=None, description="contact id"),
        params: dict | None = Field(default=None, description="params"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage contacts operations.

        Actions:
          - 'get_outlook_contact': Get Outlook contact.
          - 'create_outlook_contact': Create Outlook contact.
          - 'update_outlook_contact': Update Outlook contact.
          - 'delete_outlook_contact': Delete Outlook contact.
        """
        kwargs: dict[str, Any]
        if action == "get_outlook_contact":
            kwargs = {"contact_id": contact_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_outlook_contact(**kwargs)
        if action == "create_outlook_contact":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_outlook_contact(**kwargs)
        if action == "update_outlook_contact":
            kwargs = {"contact_id": contact_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_outlook_contact(**kwargs)
        if action == "delete_outlook_contact":
            kwargs = {"contact_id": contact_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_outlook_contact(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_outlook_contact', 'create_outlook_contact', 'update_outlook_contact', 'delete_outlook_contact")


def register_user_tools(mcp: FastMCP):
    @mcp.tool(tags={"user"})
    async def microsoft_user(
        action: str = Field(description="Action to perform. Must be one of: 'get_current_user', 'get_me'"),
        params: dict | None = Field(default=None, description="params"),
        client=Depends(get_client),
    ) -> dict:
        """Manage user operations.

        Actions:
          - 'get_current_user': Get current user (alias for get_me).
          - 'get_me': Get the current user.
        """
        kwargs: dict[str, Any]
        if action == "get_current_user":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_current_user(**kwargs)
        if action == "get_me":
            kwargs = {}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_me(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_current_user', 'get_me")


def register_chat_tools(mcp: FastMCP):
    @mcp.tool(tags={"chat"})
    async def microsoft_chat(
        action: str = Field(description="Action to perform. Must be one of: 'get_chat'"),
        chat_id: str | None = Field(default=None, description="chat id"),
        params: dict | None = Field(default=None, description="params"),
        client=Depends(get_client),
    ) -> dict:
        """Manage chat operations.

        Actions:
          - 'get_chat': Get chat.
        """
        kwargs: dict[str, Any]
        if action == "get_chat":
            kwargs = {"chat_id": chat_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_chat(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_chat")


def register_teams_tools(mcp: FastMCP):
    @mcp.tool(tags={"teams"})
    async def microsoft_teams(
        action: str = Field(description="Action to perform. Must be one of: 'get_team', 'get_team_channel'"),
        team_id: str | None = Field(default=None, description="team id"),
        params: dict | None = Field(default=None, description="params"),
        channel_id: str | None = Field(default=None, description="channel id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage teams operations.

        Actions:
          - 'get_team': Get team.
          - 'get_team_channel': Get team channel.
        """
        kwargs: dict[str, Any]
        if action == "get_team":
            kwargs = {"team_id": team_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_team(**kwargs)
        if action == "get_team_channel":
            kwargs = {"team_id": team_id, "channel_id": channel_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_team_channel(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_team', 'get_team_channel")


def register_sites_tools(mcp: FastMCP):
    @mcp.tool(tags={"sites"})
    async def microsoft_sites(
        action: str = Field(description="Action to perform. Must be one of: 'list_sites', 'get_site', 'get_sharepoint_site_by_path', 'get_sharepoint_sites_delta'"),
        params: dict | None = Field(default=None, description="params"),
        site_id: str | None = Field(default=None, description="site id"),
        path: str | None = Field(default=None, description="path"),
        client=Depends(get_client),
    ) -> dict:
        """Manage sites operations.

        Actions:
          - 'list_sites': List SharePoint sites.
          - 'get_site': Get SharePoint site.
          - 'get_sharepoint_site_by_path': Get SharePoint site by path.
          - 'get_sharepoint_sites_delta': Get SharePoint sites delta.
        """
        kwargs: dict[str, Any]
        if action == "list_sites":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_sites(**kwargs)
        if action == "get_site":
            kwargs = {"site_id": site_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_site(**kwargs)
        if action == "get_sharepoint_site_by_path":
            kwargs = {"site_id": site_id, "path": path, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sharepoint_site_by_path(**kwargs)
        if action == "get_sharepoint_sites_delta":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sharepoint_sites_delta(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_sites', 'get_site', 'get_sharepoint_site_by_path', 'get_sharepoint_sites_delta")


def register_search_tools(mcp: FastMCP):
    @mcp.tool(tags={"search"})
    async def microsoft_search(
        action: str = Field(description="Action to perform. Must be one of: 'search_query'"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        params: dict | None = Field(default=None, description="params"),
        client=Depends(get_client),
    ) -> dict:
        """Manage search operations.

        Actions:
          - 'search_query': Search query.
        """
        kwargs: dict[str, Any]
        if action == "search_query":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.search_query(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: search_query")


def register_groups_tools(mcp: FastMCP):
    @mcp.tool(tags={"groups"})
    async def microsoft_groups(
        action: str = Field(description="Action to perform. Must be one of: 'list_groups', 'get_group', 'create_group', 'update_group', 'delete_group', 'list_group_members', 'add_group_member', 'remove_group_member', 'list_group_owners', 'list_group_conversations', 'list_group_drives'"),
        params: dict | None = Field(default=None, description="params"),
        group_id: str | None = Field(default=None, description="group id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        member_id: str | None = Field(default=None, description="member id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage groups operations.

        Actions:
          - 'list_groups': List all Microsoft 365 groups and security groups.
          - 'get_group': Get a specific group.
          - 'create_group': Create a new group.
          - 'update_group': Update a group.
          - 'delete_group': Delete a group.
          - 'list_group_members': List group members.
          - 'add_group_member': Add a member to a group.
          - 'remove_group_member': Remove a member from a group.
          - 'list_group_owners': List group owners.
          - 'list_group_conversations': List group conversations.
          - 'list_group_drives': List group drives.
        """
        kwargs: dict[str, Any]
        if action == "list_groups":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_groups(**kwargs)
        if action == "get_group":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_group(**kwargs)
        if action == "create_group":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_group(**kwargs)
        if action == "update_group":
            kwargs = {"group_id": group_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_group(**kwargs)
        if action == "delete_group":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_group(**kwargs)
        if action == "list_group_members":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_group_members(**kwargs)
        if action == "add_group_member":
            kwargs = {"group_id": group_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.add_group_member(**kwargs)
        if action == "remove_group_member":
            kwargs = {"group_id": group_id, "member_id": member_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.remove_group_member(**kwargs)
        if action == "list_group_owners":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_group_owners(**kwargs)
        if action == "list_group_conversations":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_group_conversations(**kwargs)
        if action == "list_group_drives":
            kwargs = {"group_id": group_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_group_drives(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_groups', 'get_group', 'create_group', 'update_group', 'delete_group', 'list_group_members', 'add_group_member', 'remove_group_member', 'list_group_owners', 'list_group_conversations', 'list_group_drives")


def register_admin_tools(mcp: FastMCP):
    @mcp.tool(tags={"admin"})
    async def microsoft_admin(
        action: str = Field(description="Action to perform. Must be one of: 'list_service_health', 'get_service_health', 'list_service_health_issues', 'get_service_health_issue', 'list_service_update_messages', 'get_service_update_message', 'get_admin_sharepoint', 'update_admin_sharepoint', 'list_delegated_admin_relationships', 'get_delegated_admin_relationship'"),
        params: dict | None = Field(default=None, description="params"),
        service_name: str | None = Field(default=None, description="service name"),
        issue_id: str | None = Field(default=None, description="issue id"),
        message_id: str | None = Field(default=None, description="message id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        rel_id: str | None = Field(default=None, description="rel id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage admin operations.

        Actions:
          - 'list_service_health': List service health overviews.
          - 'get_service_health': Get service health for a specific service.
          - 'list_service_health_issues': List service health issues.
          - 'get_service_health_issue': Get a specific service health issue.
          - 'list_service_update_messages': List service update messages.
          - 'get_service_update_message': Get a specific service update message.
          - 'get_admin_sharepoint': Get SharePoint admin settings.
          - 'update_admin_sharepoint': Update SharePoint admin settings.
          - 'list_delegated_admin_relationships': List delegated admin relationships.
          - 'get_delegated_admin_relationship': Get a specific delegated admin relationship.
        """
        kwargs: dict[str, Any]
        if action == "list_service_health":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_service_health(**kwargs)
        if action == "get_service_health":
            kwargs = {"service_name": service_name, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_service_health(**kwargs)
        if action == "list_service_health_issues":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_service_health_issues(**kwargs)
        if action == "get_service_health_issue":
            kwargs = {"issue_id": issue_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_service_health_issue(**kwargs)
        if action == "list_service_update_messages":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_service_update_messages(**kwargs)
        if action == "get_service_update_message":
            kwargs = {"message_id": message_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_service_update_message(**kwargs)
        if action == "get_admin_sharepoint":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_admin_sharepoint(**kwargs)
        if action == "update_admin_sharepoint":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_admin_sharepoint(**kwargs)
        if action == "list_delegated_admin_relationships":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_delegated_admin_relationships(**kwargs)
        if action == "get_delegated_admin_relationship":
            kwargs = {"rel_id": rel_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_delegated_admin_relationship(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_service_health', 'get_service_health', 'list_service_health_issues', 'get_service_health_issue', 'list_service_update_messages', 'get_service_update_message', 'get_admin_sharepoint', 'update_admin_sharepoint', 'list_delegated_admin_relationships', 'get_delegated_admin_relationship")


def register_organization_tools(mcp: FastMCP):
    @mcp.tool(tags={"organization"})
    async def microsoft_organization(
        action: str = Field(description="Action to perform. Must be one of: 'list_organization', 'get_organization', 'update_organization', 'get_org_branding', 'update_org_branding'"),
        params: dict | None = Field(default=None, description="params"),
        org_id: str | None = Field(default=None, description="org id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage organization operations.

        Actions:
          - 'list_organization': List organization properties.
          - 'get_organization': Get organization by ID.
          - 'update_organization': Update organization properties.
          - 'get_org_branding': Get organization branding.
          - 'update_org_branding': Update organization branding.
        """
        kwargs: dict[str, Any]
        if action == "list_organization":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_organization(**kwargs)
        if action == "get_organization":
            kwargs = {"org_id": org_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_organization(**kwargs)
        if action == "update_organization":
            kwargs = {"org_id": org_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_organization(**kwargs)
        if action == "get_org_branding":
            kwargs = {"org_id": org_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_org_branding(**kwargs)
        if action == "update_org_branding":
            kwargs = {"org_id": org_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_org_branding(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_organization', 'get_organization', 'update_organization', 'get_org_branding', 'update_org_branding")


def register_domains_tools(mcp: FastMCP):
    @mcp.tool(tags={"domains"})
    async def microsoft_domains(
        action: str = Field(description="Action to perform. Must be one of: 'list_domains', 'get_domain', 'create_domain', 'delete_domain', 'verify_domain', 'list_domain_service_configuration_records'"),
        params: dict | None = Field(default=None, description="params"),
        domain_id: str | None = Field(default=None, description="domain id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage domains operations.

        Actions:
          - 'list_domains': List tenant domains.
          - 'get_domain': Get domain details.
          - 'create_domain': Add a domain to the tenant.
          - 'delete_domain': Delete a domain.
          - 'verify_domain': Verify domain ownership.
          - 'list_domain_service_configuration_records': List domain service configuration DNS records.
        """
        kwargs: dict[str, Any]
        if action == "list_domains":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_domains(**kwargs)
        if action == "get_domain":
            kwargs = {"domain_id": domain_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_domain(**kwargs)
        if action == "create_domain":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_domain(**kwargs)
        if action == "delete_domain":
            kwargs = {"domain_id": domain_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_domain(**kwargs)
        if action == "verify_domain":
            kwargs = {"domain_id": domain_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.verify_domain(**kwargs)
        if action == "list_domain_service_configuration_records":
            kwargs = {"domain_id": domain_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_domain_service_configuration_records(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_domains', 'get_domain', 'create_domain', 'delete_domain', 'verify_domain', 'list_domain_service_configuration_records")


def register_subscriptions_tools(mcp: FastMCP):
    @mcp.tool(tags={"subscriptions"})
    async def microsoft_subscriptions(
        action: str = Field(description="Action to perform. Must be one of: 'list_subscriptions', 'get_subscription', 'create_subscription', 'update_subscription', 'delete_subscription'"),
        params: dict | None = Field(default=None, description="params"),
        subscription_id: str | None = Field(default=None, description="subscription id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage subscriptions operations.

        Actions:
          - 'list_subscriptions': List active webhook subscriptions.
          - 'get_subscription': Get a specific subscription.
          - 'create_subscription': Create a subscription for change notifications.
          - 'update_subscription': Update/renew a subscription.
          - 'delete_subscription': Delete a subscription.
        """
        kwargs: dict[str, Any]
        if action == "list_subscriptions":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_subscriptions(**kwargs)
        if action == "get_subscription":
            kwargs = {"subscription_id": subscription_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_subscription(**kwargs)
        if action == "create_subscription":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_subscription(**kwargs)
        if action == "update_subscription":
            kwargs = {"subscription_id": subscription_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_subscription(**kwargs)
        if action == "delete_subscription":
            kwargs = {"subscription_id": subscription_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_subscription(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_subscriptions', 'get_subscription', 'create_subscription', 'update_subscription', 'delete_subscription")


def register_communications_tools(mcp: FastMCP):
    @mcp.tool(tags={"communications"})
    async def microsoft_communications(
        action: str = Field(description="Action to perform. Must be one of: 'list_online_meetings', 'get_online_meeting', 'create_online_meeting', 'update_online_meeting', 'delete_online_meeting', 'list_call_records', 'get_call_record', 'list_presences', 'get_presence', 'get_my_presence'"),
        params: dict | None = Field(default=None, description="params"),
        meeting_id: str | None = Field(default=None, description="meeting id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        call_id: str | None = Field(default=None, description="call id"),
        user_id: str | None = Field(default=None, description="user id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage communications operations.

        Actions:
          - 'list_online_meetings': List online meetings for the current user.
          - 'get_online_meeting': Get a specific online meeting.
          - 'create_online_meeting': Create a new online meeting.
          - 'update_online_meeting': Update an online meeting.
          - 'delete_online_meeting': Delete an online meeting.
          - 'list_call_records': List call records.
          - 'get_call_record': Get a specific call record.
          - 'list_presences': List presence information for users.
          - 'get_presence': Get presence for a specific user.
          - 'get_my_presence': Get current user's presence.
        """
        kwargs: dict[str, Any]
        if action == "list_online_meetings":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_online_meetings(**kwargs)
        if action == "get_online_meeting":
            kwargs = {"meeting_id": meeting_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_online_meeting(**kwargs)
        if action == "create_online_meeting":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_online_meeting(**kwargs)
        if action == "update_online_meeting":
            kwargs = {"meeting_id": meeting_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_online_meeting(**kwargs)
        if action == "delete_online_meeting":
            kwargs = {"meeting_id": meeting_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_online_meeting(**kwargs)
        if action == "list_call_records":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_call_records(**kwargs)
        if action == "get_call_record":
            kwargs = {"call_id": call_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_call_record(**kwargs)
        if action == "list_presences":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_presences(**kwargs)
        if action == "get_presence":
            kwargs = {"user_id": user_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_presence(**kwargs)
        if action == "get_my_presence":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_my_presence(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_online_meetings', 'get_online_meeting', 'create_online_meeting', 'update_online_meeting', 'delete_online_meeting', 'list_call_records', 'get_call_record', 'list_presences', 'get_presence', 'get_my_presence")


def register_identity_tools(mcp: FastMCP):
    @mcp.tool(tags={"identity"})
    async def microsoft_identity(
        action: str = Field(description="Action to perform. Must be one of: 'create_invitation', 'list_conditional_access_policies', 'get_conditional_access_policy', 'create_conditional_access_policy', 'update_conditional_access_policy', 'delete_conditional_access_policy', 'list_access_reviews', 'get_access_review', 'list_entitlement_access_packages', 'list_lifecycle_workflows'"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        params: dict | None = Field(default=None, description="params"),
        policy_id: str | None = Field(default=None, description="policy id"),
        review_id: str | None = Field(default=None, description="review id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage identity operations.

        Actions:
          - 'create_invitation': Create an invitation for a guest user.
          - 'list_conditional_access_policies': List conditional access policies.
          - 'get_conditional_access_policy': Get a specific conditional access policy.
          - 'create_conditional_access_policy': Create a conditional access policy.
          - 'update_conditional_access_policy': Update a conditional access policy.
          - 'delete_conditional_access_policy': Delete a conditional access policy.
          - 'list_access_reviews': List access review definitions.
          - 'get_access_review': Get a specific access review definition.
          - 'list_entitlement_access_packages': List entitlement management access packages.
          - 'list_lifecycle_workflows': List lifecycle management workflows.
        """
        kwargs: dict[str, Any]
        if action == "create_invitation":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_invitation(**kwargs)
        if action == "list_conditional_access_policies":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_conditional_access_policies(**kwargs)
        if action == "get_conditional_access_policy":
            kwargs = {"policy_id": policy_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_conditional_access_policy(**kwargs)
        if action == "create_conditional_access_policy":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_conditional_access_policy(**kwargs)
        if action == "update_conditional_access_policy":
            kwargs = {"policy_id": policy_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_conditional_access_policy(**kwargs)
        if action == "delete_conditional_access_policy":
            kwargs = {"policy_id": policy_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_conditional_access_policy(**kwargs)
        if action == "list_access_reviews":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_access_reviews(**kwargs)
        if action == "get_access_review":
            kwargs = {"review_id": review_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_access_review(**kwargs)
        if action == "list_entitlement_access_packages":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_entitlement_access_packages(**kwargs)
        if action == "list_lifecycle_workflows":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_lifecycle_workflows(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: create_invitation', 'list_conditional_access_policies', 'get_conditional_access_policy', 'create_conditional_access_policy', 'update_conditional_access_policy', 'delete_conditional_access_policy', 'list_access_reviews', 'get_access_review', 'list_entitlement_access_packages', 'list_lifecycle_workflows")


def register_security_tools(mcp: FastMCP):
    @mcp.tool(tags={"security"})
    async def microsoft_security(
        action: str = Field(description="Action to perform. Must be one of: 'list_security_alerts', 'get_security_alert', 'update_security_alert', 'list_security_incidents', 'get_security_incident', 'update_security_incident', 'list_secure_scores', 'list_threat_intelligence_hosts', 'get_threat_intelligence_host', 'run_hunting_query', 'list_risk_detections', 'get_risk_detection', 'list_risky_users', 'get_risky_user', 'dismiss_risky_user', 'list_sensitivity_labels', 'get_sensitivity_label'"),
        params: dict | None = Field(default=None, description="params"),
        alert_id: str | None = Field(default=None, description="alert id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        incident_id: str | None = Field(default=None, description="incident id"),
        host_id: str | None = Field(default=None, description="host id"),
        risk_id: str | None = Field(default=None, description="risk id"),
        user_id: str | None = Field(default=None, description="user id"),
        label_id: str | None = Field(default=None, description="label id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage security operations.

        Actions:
          - 'list_security_alerts': List security alerts (v2).
          - 'get_security_alert': Get a specific security alert.
          - 'update_security_alert': Update a security alert (e.g. change status, assign).
          - 'list_security_incidents': List security incidents.
          - 'get_security_incident': Get a specific security incident.
          - 'update_security_incident': Update a security incident.
          - 'list_secure_scores': List secure scores.
          - 'list_threat_intelligence_hosts': List threat intelligence hosts.
          - 'get_threat_intelligence_host': Get a specific threat intelligence host.
          - 'run_hunting_query': Run an advanced hunting query.
          - 'list_risk_detections': List risk detections.
          - 'get_risk_detection': Get a specific risk detection.
          - 'list_risky_users': List risky users.
          - 'get_risky_user': Get a specific risky user.
          - 'dismiss_risky_user': Dismiss a risky user.
          - 'list_sensitivity_labels': List sensitivity labels.
          - 'get_sensitivity_label': Get a specific sensitivity label.
        """
        kwargs: dict[str, Any]
        if action == "list_security_alerts":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_security_alerts(**kwargs)
        if action == "get_security_alert":
            kwargs = {"alert_id": alert_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_security_alert(**kwargs)
        if action == "update_security_alert":
            kwargs = {"alert_id": alert_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_security_alert(**kwargs)
        if action == "list_security_incidents":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_security_incidents(**kwargs)
        if action == "get_security_incident":
            kwargs = {"incident_id": incident_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_security_incident(**kwargs)
        if action == "update_security_incident":
            kwargs = {"incident_id": incident_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_security_incident(**kwargs)
        if action == "list_secure_scores":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_secure_scores(**kwargs)
        if action == "list_threat_intelligence_hosts":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_threat_intelligence_hosts(**kwargs)
        if action == "get_threat_intelligence_host":
            kwargs = {"host_id": host_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_threat_intelligence_host(**kwargs)
        if action == "run_hunting_query":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.run_hunting_query(**kwargs)
        if action == "list_risk_detections":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_risk_detections(**kwargs)
        if action == "get_risk_detection":
            kwargs = {"risk_id": risk_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_risk_detection(**kwargs)
        if action == "list_risky_users":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_risky_users(**kwargs)
        if action == "get_risky_user":
            kwargs = {"user_id": user_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_risky_user(**kwargs)
        if action == "dismiss_risky_user":
            kwargs = {"user_id": user_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.dismiss_risky_user(**kwargs)
        if action == "list_sensitivity_labels":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_sensitivity_labels(**kwargs)
        if action == "get_sensitivity_label":
            kwargs = {"label_id": label_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sensitivity_label(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_security_alerts', 'get_security_alert', 'update_security_alert', 'list_security_incidents', 'get_security_incident', 'update_security_incident', 'list_secure_scores', 'list_threat_intelligence_hosts', 'get_threat_intelligence_host', 'run_hunting_query', 'list_risk_detections', 'get_risk_detection', 'list_risky_users', 'get_risky_user', 'dismiss_risky_user', 'list_sensitivity_labels', 'get_sensitivity_label")


def register_audit_tools(mcp: FastMCP):
    @mcp.tool(tags={"audit"})
    async def microsoft_audit(
        action: str = Field(description="Action to perform. Must be one of: 'list_directory_audits', 'get_directory_audit', 'list_sign_in_logs', 'get_sign_in_log', 'list_provisioning_logs'"),
        params: dict | None = Field(default=None, description="params"),
        audit_id: str | None = Field(default=None, description="audit id"),
        sign_in_id: str | None = Field(default=None, description="sign in id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage audit operations.

        Actions:
          - 'list_directory_audits': List directory audit logs.
          - 'get_directory_audit': Get a specific directory audit entry.
          - 'list_sign_in_logs': List sign-in logs.
          - 'get_sign_in_log': Get a specific sign-in log entry.
          - 'list_provisioning_logs': List provisioning logs.
        """
        kwargs: dict[str, Any]
        if action == "list_directory_audits":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_directory_audits(**kwargs)
        if action == "get_directory_audit":
            kwargs = {"audit_id": audit_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_directory_audit(**kwargs)
        if action == "list_sign_in_logs":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_sign_in_logs(**kwargs)
        if action == "get_sign_in_log":
            kwargs = {"sign_in_id": sign_in_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sign_in_log(**kwargs)
        if action == "list_provisioning_logs":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_provisioning_logs(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_directory_audits', 'get_directory_audit', 'list_sign_in_logs', 'get_sign_in_log', 'list_provisioning_logs")


def register_reports_tools(mcp: FastMCP):
    @mcp.tool(tags={"reports"})
    async def microsoft_reports(
        action: str = Field(description="Action to perform. Must be one of: 'get_email_activity_report', 'get_mailbox_usage_report', 'get_office365_active_users', 'get_sharepoint_activity_report', 'get_teams_user_activity', 'get_onedrive_usage_report'"),
        period: str | None = Field(default=None, description="period"),
        params: dict | None = Field(default=None, description="params"),
        client=Depends(get_client),
    ) -> dict:
        """Manage reports operations.

        Actions:
          - 'get_email_activity_report': Get email activity user detail report.
          - 'get_mailbox_usage_report': Get mailbox usage detail report.
          - 'get_office365_active_users': Get Office 365 active user detail report.
          - 'get_sharepoint_activity_report': Get SharePoint activity user detail report.
          - 'get_teams_user_activity': Get Teams user activity detail report.
          - 'get_onedrive_usage_report': Get OneDrive usage account detail report.
        """
        kwargs: dict[str, Any]
        if action == "get_email_activity_report":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_email_activity_report(**kwargs)
        if action == "get_mailbox_usage_report":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_mailbox_usage_report(**kwargs)
        if action == "get_office365_active_users":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_office365_active_users(**kwargs)
        if action == "get_sharepoint_activity_report":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_sharepoint_activity_report(**kwargs)
        if action == "get_teams_user_activity":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_teams_user_activity(**kwargs)
        if action == "get_onedrive_usage_report":
            kwargs = {"period": period, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_onedrive_usage_report(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_email_activity_report', 'get_mailbox_usage_report', 'get_office365_active_users', 'get_sharepoint_activity_report', 'get_teams_user_activity', 'get_onedrive_usage_report")


def register_applications_tools(mcp: FastMCP):
    @mcp.tool(tags={"applications"})
    async def microsoft_applications(
        action: str = Field(description="Action to perform. Must be one of: 'list_applications', 'get_application', 'create_application', 'update_application', 'delete_application', 'add_application_password', 'remove_application_password', 'list_service_principals', 'get_service_principal', 'create_service_principal', 'update_service_principal', 'delete_service_principal'"),
        params: dict | None = Field(default=None, description="params"),
        app_id: str | None = Field(default=None, description="app id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        sp_id: str | None = Field(default=None, description="sp id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage applications operations.

        Actions:
          - 'list_applications': List app registrations.
          - 'get_application': Get a specific application.
          - 'create_application': Create an application registration.
          - 'update_application': Update an application.
          - 'delete_application': Delete an application.
          - 'add_application_password': Add a password credential to an application.
          - 'remove_application_password': Remove a password credential from an application.
          - 'list_service_principals': List service principals.
          - 'get_service_principal': Get a specific service principal.
          - 'create_service_principal': Create a service principal.
          - 'update_service_principal': Update a service principal.
          - 'delete_service_principal': Delete a service principal.
        """
        kwargs: dict[str, Any]
        if action == "list_applications":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_applications(**kwargs)
        if action == "get_application":
            kwargs = {"app_id": app_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_application(**kwargs)
        if action == "create_application":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_application(**kwargs)
        if action == "update_application":
            kwargs = {"app_id": app_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_application(**kwargs)
        if action == "delete_application":
            kwargs = {"app_id": app_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_application(**kwargs)
        if action == "add_application_password":
            kwargs = {"app_id": app_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.add_application_password(**kwargs)
        if action == "remove_application_password":
            kwargs = {"app_id": app_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.remove_application_password(**kwargs)
        if action == "list_service_principals":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_service_principals(**kwargs)
        if action == "get_service_principal":
            kwargs = {"sp_id": sp_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_service_principal(**kwargs)
        if action == "create_service_principal":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_service_principal(**kwargs)
        if action == "update_service_principal":
            kwargs = {"sp_id": sp_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_service_principal(**kwargs)
        if action == "delete_service_principal":
            kwargs = {"sp_id": sp_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_service_principal(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_applications', 'get_application', 'create_application', 'update_application', 'delete_application', 'add_application_password', 'remove_application_password', 'list_service_principals', 'get_service_principal', 'create_service_principal', 'update_service_principal', 'delete_service_principal")


def register_directory_tools(mcp: FastMCP):
    @mcp.tool(tags={"directory"})
    async def microsoft_directory(
        action: str = Field(description="Action to perform. Must be one of: 'list_directory_objects', 'get_directory_object', 'list_directory_roles', 'get_directory_role', 'list_directory_role_templates', 'list_deleted_items', 'restore_deleted_item', 'list_role_definitions', 'get_role_definition', 'list_role_assignments', 'get_role_assignment', 'create_role_assignment'"),
        params: dict | None = Field(default=None, description="params"),
        object_id: str | None = Field(default=None, description="object id"),
        role_id: str | None = Field(default=None, description="role id"),
        assignment_id: str | None = Field(default=None, description="assignment id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage directory operations.

        Actions:
          - 'list_directory_objects': List directory objects.
          - 'get_directory_object': Get a specific directory object.
          - 'list_directory_roles': List directory roles.
          - 'get_directory_role': Get a specific directory role.
          - 'list_directory_role_templates': List directory role templates.
          - 'list_deleted_items': List deleted directory items.
          - 'restore_deleted_item': Restore a deleted directory item.
          - 'list_role_definitions': List role definitions.
          - 'get_role_definition': Get a specific role definition.
          - 'list_role_assignments': List role assignments.
          - 'get_role_assignment': Get a specific role assignment.
          - 'create_role_assignment': Create a role assignment.
        """
        kwargs: dict[str, Any]
        if action == "list_directory_objects":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_directory_objects(**kwargs)
        if action == "get_directory_object":
            kwargs = {"object_id": object_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_directory_object(**kwargs)
        if action == "list_directory_roles":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_directory_roles(**kwargs)
        if action == "get_directory_role":
            kwargs = {"role_id": role_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_directory_role(**kwargs)
        if action == "list_directory_role_templates":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_directory_role_templates(**kwargs)
        if action == "list_deleted_items":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_deleted_items(**kwargs)
        if action == "restore_deleted_item":
            kwargs = {"object_id": object_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.restore_deleted_item(**kwargs)
        if action == "list_role_definitions":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_role_definitions(**kwargs)
        if action == "get_role_definition":
            kwargs = {"role_id": role_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_role_definition(**kwargs)
        if action == "list_role_assignments":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_role_assignments(**kwargs)
        if action == "get_role_assignment":
            kwargs = {"assignment_id": assignment_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_role_assignment(**kwargs)
        if action == "create_role_assignment":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_role_assignment(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_directory_objects', 'get_directory_object', 'list_directory_roles', 'get_directory_role', 'list_directory_role_templates', 'list_deleted_items', 'restore_deleted_item', 'list_role_definitions', 'get_role_definition', 'list_role_assignments', 'get_role_assignment', 'create_role_assignment")


def register_policies_tools(mcp: FastMCP):
    @mcp.tool(tags={"policies"})
    async def microsoft_policies(
        action: str = Field(description="Action to perform. Must be one of: 'get_authorization_policy', 'list_token_lifetime_policies', 'list_token_issuance_policies', 'list_permission_grant_policies', 'get_admin_consent_policy'"),
        params: dict | None = Field(default=None, description="params"),
        client=Depends(get_client),
    ) -> dict:
        """Manage policies operations.

        Actions:
          - 'get_authorization_policy': Get the authorization policy.
          - 'list_token_lifetime_policies': List token lifetime policies.
          - 'list_token_issuance_policies': List token issuance policies.
          - 'list_permission_grant_policies': List permission grant policies.
          - 'get_admin_consent_policy': Get the admin consent request policy.
        """
        kwargs: dict[str, Any]
        if action == "get_authorization_policy":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_authorization_policy(**kwargs)
        if action == "list_token_lifetime_policies":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_token_lifetime_policies(**kwargs)
        if action == "list_token_issuance_policies":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_token_issuance_policies(**kwargs)
        if action == "list_permission_grant_policies":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_permission_grant_policies(**kwargs)
        if action == "get_admin_consent_policy":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_admin_consent_policy(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: get_authorization_policy', 'list_token_lifetime_policies', 'list_token_issuance_policies', 'list_permission_grant_policies', 'get_admin_consent_policy")


def register_devices_tools(mcp: FastMCP):
    @mcp.tool(tags={"devices"})
    async def microsoft_devices(
        action: str = Field(description="Action to perform. Must be one of: 'list_devices', 'get_device', 'delete_device', 'list_managed_devices', 'get_managed_device', 'list_device_compliance_policies', 'list_device_configurations', 'wipe_managed_device', 'retire_managed_device'"),
        params: dict | None = Field(default=None, description="params"),
        device_id: str | None = Field(default=None, description="device id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage devices operations.

        Actions:
          - 'list_devices': List devices registered in the directory.
          - 'get_device': Get a specific device.
          - 'delete_device': Delete a device.
          - 'list_managed_devices': List managed devices.
          - 'get_managed_device': Get a specific managed device.
          - 'list_device_compliance_policies': List device compliance policies.
          - 'list_device_configurations': List device configurations.
          - 'wipe_managed_device': Wipe a managed device.
          - 'retire_managed_device': Retire a managed device.
        """
        kwargs: dict[str, Any]
        if action == "list_devices":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_devices(**kwargs)
        if action == "get_device":
            kwargs = {"device_id": device_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_device(**kwargs)
        if action == "delete_device":
            kwargs = {"device_id": device_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_device(**kwargs)
        if action == "list_managed_devices":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_managed_devices(**kwargs)
        if action == "get_managed_device":
            kwargs = {"device_id": device_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_managed_device(**kwargs)
        if action == "list_device_compliance_policies":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_device_compliance_policies(**kwargs)
        if action == "list_device_configurations":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_device_configurations(**kwargs)
        if action == "wipe_managed_device":
            kwargs = {"device_id": device_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.wipe_managed_device(**kwargs)
        if action == "retire_managed_device":
            kwargs = {"device_id": device_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.retire_managed_device(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_devices', 'get_device', 'delete_device', 'list_managed_devices', 'get_managed_device', 'list_device_compliance_policies', 'list_device_configurations', 'wipe_managed_device', 'retire_managed_device")


def register_education_tools(mcp: FastMCP):
    @mcp.tool(tags={"education"})
    async def microsoft_education(
        action: str = Field(description="Action to perform. Must be one of: 'list_education_classes', 'get_education_class', 'list_education_schools', 'get_education_school', 'list_education_users', 'list_education_assignments'"),
        params: dict | None = Field(default=None, description="params"),
        class_id: str | None = Field(default=None, description="class id"),
        school_id: str | None = Field(default=None, description="school id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage education operations.

        Actions:
          - 'list_education_classes': List education classes.
          - 'get_education_class': Get a specific education class.
          - 'list_education_schools': List education schools.
          - 'get_education_school': Get a specific education school.
          - 'list_education_users': List education users.
          - 'list_education_assignments': List assignments for an education class.
        """
        kwargs: dict[str, Any]
        if action == "list_education_classes":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_education_classes(**kwargs)
        if action == "get_education_class":
            kwargs = {"class_id": class_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_education_class(**kwargs)
        if action == "list_education_schools":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_education_schools(**kwargs)
        if action == "get_education_school":
            kwargs = {"school_id": school_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_education_school(**kwargs)
        if action == "list_education_users":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_education_users(**kwargs)
        if action == "list_education_assignments":
            kwargs = {"class_id": class_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_education_assignments(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_education_classes', 'get_education_class', 'list_education_schools', 'get_education_school', 'list_education_users', 'list_education_assignments")


def register_agreements_tools(mcp: FastMCP):
    @mcp.tool(tags={"agreements"})
    async def microsoft_agreements(
        action: str = Field(description="Action to perform. Must be one of: 'list_agreements', 'get_agreement', 'create_agreement', 'delete_agreement'"),
        params: dict | None = Field(default=None, description="params"),
        agreement_id: str | None = Field(default=None, description="agreement id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage agreements operations.

        Actions:
          - 'list_agreements': List agreements (terms of use).
          - 'get_agreement': Get a specific agreement.
          - 'create_agreement': Create an agreement.
          - 'delete_agreement': Delete an agreement.
        """
        kwargs: dict[str, Any]
        if action == "list_agreements":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_agreements(**kwargs)
        if action == "get_agreement":
            kwargs = {"agreement_id": agreement_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_agreement(**kwargs)
        if action == "create_agreement":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_agreement(**kwargs)
        if action == "delete_agreement":
            kwargs = {"agreement_id": agreement_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_agreement(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_agreements', 'get_agreement', 'create_agreement', 'delete_agreement")


def register_places_tools(mcp: FastMCP):
    @mcp.tool(tags={"places"})
    async def microsoft_places(
        action: str = Field(description="Action to perform. Must be one of: 'list_rooms', 'list_room_lists', 'get_place', 'update_place'"),
        params: dict | None = Field(default=None, description="params"),
        place_id: str | None = Field(default=None, description="place id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage places operations.

        Actions:
          - 'list_rooms': List rooms.
          - 'list_room_lists': List room lists.
          - 'get_place': Get a specific place.
          - 'update_place': Update a place.
        """
        kwargs: dict[str, Any]
        if action == "list_rooms":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_rooms(**kwargs)
        if action == "list_room_lists":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_room_lists(**kwargs)
        if action == "get_place":
            kwargs = {"place_id": place_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_place(**kwargs)
        if action == "update_place":
            kwargs = {"place_id": place_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.update_place(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_rooms', 'list_room_lists', 'get_place', 'update_place")


def register_print_tools(mcp: FastMCP):
    @mcp.tool(tags={"print"})
    async def microsoft_print(
        action: str = Field(description="Action to perform. Must be one of: 'list_printers', 'get_printer', 'list_print_jobs', 'create_print_job', 'list_print_shares'"),
        params: dict | None = Field(default=None, description="params"),
        printer_id: str | None = Field(default=None, description="printer id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage print operations.

        Actions:
          - 'list_printers': List printers.
          - 'get_printer': Get a specific printer.
          - 'list_print_jobs': List print jobs for a printer.
          - 'create_print_job': Create a print job.
          - 'list_print_shares': List print shares.
        """
        kwargs: dict[str, Any]
        if action == "list_printers":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_printers(**kwargs)
        if action == "get_printer":
            kwargs = {"printer_id": printer_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_printer(**kwargs)
        if action == "list_print_jobs":
            kwargs = {"printer_id": printer_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_print_jobs(**kwargs)
        if action == "create_print_job":
            kwargs = {"printer_id": printer_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_print_job(**kwargs)
        if action == "list_print_shares":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_print_shares(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_printers', 'get_printer', 'list_print_jobs', 'create_print_job', 'list_print_shares")


def register_privacy_tools(mcp: FastMCP):
    @mcp.tool(tags={"privacy"})
    async def microsoft_privacy(
        action: str = Field(description="Action to perform. Must be one of: 'list_subject_rights_requests', 'get_subject_rights_request', 'create_subject_rights_request'"),
        params: dict | None = Field(default=None, description="params"),
        request_id: str | None = Field(default=None, description="request id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage privacy operations.

        Actions:
          - 'list_subject_rights_requests': List subject rights requests.
          - 'get_subject_rights_request': Get a specific subject rights request.
          - 'create_subject_rights_request': Create a subject rights request.
        """
        kwargs: dict[str, Any]
        if action == "list_subject_rights_requests":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_subject_rights_requests(**kwargs)
        if action == "get_subject_rights_request":
            kwargs = {"request_id": request_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_subject_rights_request(**kwargs)
        if action == "create_subject_rights_request":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_subject_rights_request(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_subject_rights_requests', 'get_subject_rights_request', 'create_subject_rights_request")


def register_solutions_tools(mcp: FastMCP):
    @mcp.tool(tags={"solutions"})
    async def microsoft_solutions(
        action: str = Field(description="Action to perform. Must be one of: 'list_booking_businesses', 'get_booking_business', 'list_booking_appointments', 'create_booking_appointment', 'list_virtual_events'"),
        params: dict | None = Field(default=None, description="params"),
        business_id: str | None = Field(default=None, description="business id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage solutions operations.

        Actions:
          - 'list_booking_businesses': List booking businesses.
          - 'get_booking_business': Get a specific booking business.
          - 'list_booking_appointments': List booking appointments for a business.
          - 'create_booking_appointment': Create a booking appointment.
          - 'list_virtual_events': List virtual event townhalls.
        """
        kwargs: dict[str, Any]
        if action == "list_booking_businesses":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_booking_businesses(**kwargs)
        if action == "get_booking_business":
            kwargs = {"business_id": business_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_booking_business(**kwargs)
        if action == "list_booking_appointments":
            kwargs = {"business_id": business_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_booking_appointments(**kwargs)
        if action == "create_booking_appointment":
            kwargs = {"business_id": business_id, "data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_booking_appointment(**kwargs)
        if action == "list_virtual_events":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_virtual_events(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_booking_businesses', 'get_booking_business', 'list_booking_appointments', 'create_booking_appointment', 'list_virtual_events")


def register_storage_tools(mcp: FastMCP):
    @mcp.tool(tags={"storage"})
    async def microsoft_storage(
        action: str = Field(description="Action to perform. Must be one of: 'list_file_storage_containers', 'get_file_storage_container', 'create_file_storage_container'"),
        params: dict | None = Field(default=None, description="params"),
        container_id: str | None = Field(default=None, description="container id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage storage operations.

        Actions:
          - 'list_file_storage_containers': List file storage containers.
          - 'get_file_storage_container': Get a specific file storage container.
          - 'create_file_storage_container': Create a file storage container.
        """
        kwargs: dict[str, Any]
        if action == "list_file_storage_containers":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_file_storage_containers(**kwargs)
        if action == "get_file_storage_container":
            kwargs = {"container_id": container_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_file_storage_container(**kwargs)
        if action == "create_file_storage_container":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_file_storage_container(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_file_storage_containers', 'get_file_storage_container', 'create_file_storage_container")


def register_employee_experience_tools(mcp: FastMCP):
    @mcp.tool(tags={"employee_experience"})
    async def microsoft_employee_experience(
        action: str = Field(description="Action to perform. Must be one of: 'list_learning_providers', 'get_learning_provider', 'list_learning_course_activities'"),
        params: dict | None = Field(default=None, description="params"),
        provider_id: str | None = Field(default=None, description="provider id"),
        client=Depends(get_client),
    ) -> dict:
        """Manage employee experience operations.

        Actions:
          - 'list_learning_providers': List learning providers.
          - 'get_learning_provider': Get a specific learning provider.
          - 'list_learning_course_activities': List learning course activities for the current user.
        """
        kwargs: dict[str, Any]
        if action == "list_learning_providers":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_learning_providers(**kwargs)
        if action == "get_learning_provider":
            kwargs = {"provider_id": provider_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_learning_provider(**kwargs)
        if action == "list_learning_course_activities":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_learning_course_activities(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_learning_providers', 'get_learning_provider', 'list_learning_course_activities")


def register_connections_tools(mcp: FastMCP):
    @mcp.tool(tags={"connections"})
    async def microsoft_connections(
        action: str = Field(description="Action to perform. Must be one of: 'list_external_connections', 'get_external_connection', 'create_external_connection', 'delete_external_connection'"),
        params: dict | None = Field(default=None, description="params"),
        connection_id: str | None = Field(default=None, description="connection id"),
        data: dict[str, Any] | None = Field(default=None, description="data"),
        client=Depends(get_client),
    ) -> dict:
        """Manage connections operations.

        Actions:
          - 'list_external_connections': List external connections.
          - 'get_external_connection': Get a specific external connection.
          - 'create_external_connection': Create an external connection.
          - 'delete_external_connection': Delete an external connection.
        """
        kwargs: dict[str, Any]
        if action == "list_external_connections":
            kwargs = {"params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.list_external_connections(**kwargs)
        if action == "get_external_connection":
            kwargs = {"connection_id": connection_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.get_external_connection(**kwargs)
        if action == "create_external_connection":
            kwargs = {"data": data, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.create_external_connection(**kwargs)
        if action == "delete_external_connection":
            kwargs = {"connection_id": connection_id, "params": params}
            kwargs = {k: v for k, v in kwargs.items() if v is not None}
            return client.delete_external_connection(**kwargs)
        raise ValueError(f"Unknown action: {action}. Must be one of: list_external_connections', 'get_external_connection', 'create_external_connection', 'delete_external_connection")



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
    DEFAULT_EMPLOYEE_EXPERIENCETOOL = to_boolean(os.getenv("EMPLOYEE_EXPERIENCETOOL", "True"))
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
