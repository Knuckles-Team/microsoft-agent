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
import sys
from typing import Any

from agent_utilities.mcp_utilities import (
    create_mcp_server,
    load_config,
    register_tool_surface,
    resolve_action,
    run_blocking,
)
from starlette.requests import Request
from starlette.responses import JSONResponse

from microsoft_agent.api_client import MicrosoftGraphApi
from microsoft_agent.auth import get_client

__version__ = "1.0.1"

logger = get_logger(name="microsoft-agent")
logger.setLevel(logging.INFO)

_AUTH_ACTIONS = ("login", "logout", "verify_login", "list_accounts")
_META_ACTIONS = ("searches",)
_MAIL_ACTIONS = (
    "list_mail_messages",
    "list_mail_folders",
    "list_mail_folder_messages",
    "get_mail_message",
    "send_mail",
    "list_shared_mailbox_messages",
    "list_shared_mailbox_folder_messages",
    "get_shared_mailbox_message",
    "send_shared_mailbox_mail",
    "create_draft_email",
    "delete_mail_message",
    "move_mail_message",
    "update_mail_message",
    "add_mail_attachment",
    "list_mail_attachments",
    "get_mail_attachment",
    "delete_mail_attachment",
    "get_root_folder",
    "list_folder_files",
    "list_chat_messages",
    "get_chat_message",
    "send_chat_message",
    "list_channel_messages",
    "get_channel_message",
    "send_channel_message",
    "list_chat_message_replies",
    "reply_to_chat_message",
)
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
_CALENDAR_ACTIONS = (
    "list_calendar_events",
    "get_calendar_event",
    "create_calendar_event",
    "update_calendar_event",
    "delete_calendar_event",
    "list_specific_calendar_events",
    "get_specific_calendar_event",
    "create_specific_calendar_event",
    "update_specific_calendar_event",
    "delete_specific_calendar_event",
    "get_calendar_view",
    "list_calendars",
    "find_meeting_times",
)
_NOTES_ACTIONS = ("get_onenote_page_content", "create_onenote_page")
_TASKS_ACTIONS = (
    "get_todo_task",
    "create_todo_task",
    "update_todo_task",
    "delete_todo_task",
    "get_planner_plan",
    "get_planner_task",
    "create_planner_task",
    "update_planner_task",
    "update_planner_task_details",
)
_CONTACTS_ACTIONS = (
    "get_outlook_contact",
    "create_outlook_contact",
    "update_outlook_contact",
    "delete_outlook_contact",
)
_USER_ACTIONS = ("get_current_user", "get_me")
_CHAT_ACTIONS = ("get_chat",)
_TEAMS_ACTIONS = ("get_team", "get_team_channel")
_SITES_ACTIONS = (
    "list_sites",
    "get_site",
    "get_sharepoint_site_by_path",
    "get_sharepoint_sites_delta",
)
_SEARCH_ACTIONS = ("search_query", "search_tools")
_GROUPS_ACTIONS = (
    "list_groups",
    "get_group",
    "create_group",
    "update_group",
    "delete_group",
    "list_group_members",
    "add_group_member",
    "remove_group_member",
    "list_group_owners",
    "list_group_conversations",
    "list_group_drives",
)
_ADMIN_ACTIONS = (
    "list_service_health",
    "get_service_health",
    "list_service_health_issues",
    "get_service_health_issue",
    "list_service_update_messages",
    "get_service_update_message",
    "get_admin_sharepoint",
    "update_admin_sharepoint",
    "list_delegated_admin_relationships",
    "get_delegated_admin_relationship",
)
_ORGANIZATION_ACTIONS = (
    "list_organization",
    "get_organization",
    "update_organization",
    "get_org_branding",
    "update_org_branding",
)
_DOMAINS_ACTIONS = (
    "list_domains",
    "get_domain",
    "create_domain",
    "delete_domain",
    "verify_domain",
    "list_domain_service_configuration_records",
)
_SUBSCRIPTIONS_ACTIONS = (
    "list_subscriptions",
    "get_subscription",
    "create_subscription",
    "update_subscription",
    "delete_subscription",
)
_COMMUNICATIONS_ACTIONS = (
    "list_online_meetings",
    "get_online_meeting",
    "create_online_meeting",
    "update_online_meeting",
    "delete_online_meeting",
    "list_call_records",
    "get_call_record",
    "list_presences",
    "get_presence",
    "get_my_presence",
)
_IDENTITY_ACTIONS = (
    "create_invitation",
    "list_conditional_access_policies",
    "get_conditional_access_policy",
    "create_conditional_access_policy",
    "update_conditional_access_policy",
    "delete_conditional_access_policy",
    "list_access_reviews",
    "get_access_review",
    "list_entitlement_access_packages",
    "list_lifecycle_workflows",
)
_SECURITY_ACTIONS = (
    "list_security_alerts",
    "get_security_alert",
    "update_security_alert",
    "list_security_incidents",
    "get_security_incident",
    "update_security_incident",
    "list_secure_scores",
    "list_threat_intelligence_hosts",
    "get_threat_intelligence_host",
    "run_hunting_query",
    "list_risk_detections",
    "get_risk_detection",
    "list_risky_users",
    "get_risky_user",
    "dismiss_risky_user",
    "list_sensitivity_labels",
    "get_sensitivity_label",
)
_AUDIT_ACTIONS = (
    "list_directory_audits",
    "get_directory_audit",
    "list_sign_in_logs",
    "get_sign_in_log",
    "list_provisioning_logs",
)
_REPORTS_ACTIONS = (
    "get_email_activity_report",
    "get_mailbox_usage_report",
    "get_office365_active_users",
    "get_sharepoint_activity_report",
    "get_teams_user_activity",
    "get_onedrive_usage_report",
)
_APPLICATIONS_ACTIONS = (
    "list_applications",
    "get_application",
    "create_application",
    "update_application",
    "delete_application",
    "add_application_password",
    "remove_application_password",
    "list_service_principals",
    "get_service_principal",
    "create_service_principal",
    "update_service_principal",
    "delete_service_principal",
)
_DIRECTORY_ACTIONS = (
    "list_directory_objects",
    "get_directory_object",
    "list_directory_roles",
    "get_directory_role",
    "list_directory_role_templates",
    "list_deleted_items",
    "restore_deleted_item",
    "list_role_definitions",
    "get_role_definition",
    "list_role_assignments",
    "get_role_assignment",
    "create_role_assignment",
)
_POLICIES_ACTIONS = (
    "get_authorization_policy",
    "list_token_lifetime_policies",
    "list_token_issuance_policies",
    "list_permission_grant_policies",
    "get_admin_consent_policy",
)
_DEVICES_ACTIONS = (
    "list_devices",
    "get_device",
    "delete_device",
    "list_managed_devices",
    "get_managed_device",
    "list_device_compliance_policies",
    "list_device_configurations",
    "wipe_managed_device",
    "retire_managed_device",
)
_EDUCATION_ACTIONS = (
    "list_education_classes",
    "get_education_class",
    "list_education_schools",
    "get_education_school",
    "list_education_users",
    "list_education_assignments",
)
_AGREEMENTS_ACTIONS = (
    "list_agreements",
    "get_agreement",
    "create_agreement",
    "delete_agreement",
)
_PLACES_ACTIONS = ("list_rooms", "list_room_lists", "get_place", "update_place")
_PRINT_ACTIONS = (
    "list_printers",
    "get_printer",
    "list_print_jobs",
    "create_print_job",
    "list_print_shares",
)
_PRIVACY_ACTIONS = (
    "list_subject_rights_requests",
    "get_subject_rights_request",
    "create_subject_rights_request",
)
_SOLUTIONS_ACTIONS = (
    "list_booking_businesses",
    "get_booking_business",
    "list_booking_appointments",
    "create_booking_appointment",
    "list_virtual_events",
)
_STORAGE_ACTIONS = (
    "list_file_storage_containers",
    "get_file_storage_container",
    "create_file_storage_container",
)
_EMPLOYEE_EXPERIENCE_ACTIONS = (
    "list_learning_providers",
    "get_learning_provider",
    "list_learning_course_activities",
)
_CONNECTIONS_ACTIONS = (
    "list_external_connections",
    "get_external_connection",
    "create_external_connection",
    "delete_external_connection",
)


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

        resolved = resolve_action(action, _AUTH_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "login":
            return await run_blocking(client.login, **kwargs)
        if action == "logout":
            return await run_blocking(client.logout, **kwargs)
        if action == "verify_login":
            return await run_blocking(client.verify_login, **kwargs)
        if action == "list_accounts":
            return await run_blocking(client.list_accounts, **kwargs)
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

        resolved = resolve_action(action, _META_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "searches":
            return await run_blocking(client.searches, **kwargs)
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

        resolved = resolve_action(action, _MAIL_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_mail_messages":
            return await run_blocking(client.list_mail_messages, **kwargs)
        if action == "list_mail_folders":
            return await run_blocking(client.list_mail_folders, **kwargs)
        if action == "list_mail_folder_messages":
            return await run_blocking(client.list_mail_folder_messages, **kwargs)
        if action == "get_mail_message":
            return await run_blocking(client.get_mail_message, **kwargs)
        if action == "send_mail":
            return await run_blocking(client.send_mail, **kwargs)
        if action == "list_shared_mailbox_messages":
            return await run_blocking(client.list_shared_mailbox_messages, **kwargs)
        if action == "list_shared_mailbox_folder_messages":
            return await run_blocking(
                client.list_shared_mailbox_folder_messages, **kwargs
            )
        if action == "get_shared_mailbox_message":
            return await run_blocking(client.get_shared_mailbox_message, **kwargs)
        if action == "send_shared_mailbox_mail":
            return await run_blocking(client.send_shared_mailbox_mail, **kwargs)
        if action == "create_draft_email":
            return await run_blocking(client.create_draft_email, **kwargs)
        if action == "delete_mail_message":
            return await run_blocking(client.delete_mail_message, **kwargs)
        if action == "move_mail_message":
            return await run_blocking(client.move_mail_message, **kwargs)
        if action == "update_mail_message":
            return await run_blocking(client.update_mail_message, **kwargs)
        if action == "add_mail_attachment":
            return await run_blocking(client.add_mail_attachment, **kwargs)
        if action == "list_mail_attachments":
            return await run_blocking(client.list_mail_attachments, **kwargs)
        if action == "get_mail_attachment":
            return await run_blocking(client.get_mail_attachment, **kwargs)
        if action == "delete_mail_attachment":
            return await run_blocking(client.delete_mail_attachment, **kwargs)
        if action == "get_root_folder":
            return await run_blocking(client.get_root_folder, **kwargs)
        if action == "list_folder_files":
            return await run_blocking(client.list_folder_files, **kwargs)
        if action == "list_chat_messages":
            return await run_blocking(client.list_chat_messages, **kwargs)
        if action == "get_chat_message":
            return await run_blocking(client.get_chat_message, **kwargs)
        if action == "send_chat_message":
            return await run_blocking(client.send_chat_message, **kwargs)
        if action == "list_channel_messages":
            return await run_blocking(client.list_channel_messages, **kwargs)
        if action == "get_channel_message":
            return await run_blocking(client.get_channel_message, **kwargs)
        if action == "send_channel_message":
            return await run_blocking(client.send_channel_message, **kwargs)
        if action == "list_chat_message_replies":
            return await run_blocking(client.list_chat_message_replies, **kwargs)
        if action == "reply_to_chat_message":
            return await run_blocking(client.reply_to_chat_message, **kwargs)
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

        resolved = resolve_action(action, _FILES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

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

        resolved = resolve_action(action, _CALENDAR_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_calendar_events":
            return await run_blocking(client.list_calendar_events, **kwargs)
        if action == "get_calendar_event":
            return await run_blocking(client.get_calendar_event, **kwargs)
        if action == "create_calendar_event":
            return await run_blocking(client.create_calendar_event, **kwargs)
        if action == "update_calendar_event":
            return await run_blocking(client.update_calendar_event, **kwargs)
        if action == "delete_calendar_event":
            return await run_blocking(client.delete_calendar_event, **kwargs)
        if action == "list_specific_calendar_events":
            return await run_blocking(client.list_specific_calendar_events, **kwargs)
        if action == "get_specific_calendar_event":
            return await run_blocking(client.get_specific_calendar_event, **kwargs)
        if action == "create_specific_calendar_event":
            return await run_blocking(client.create_specific_calendar_event, **kwargs)
        if action == "update_specific_calendar_event":
            return await run_blocking(client.update_specific_calendar_event, **kwargs)
        if action == "delete_specific_calendar_event":
            return await run_blocking(client.delete_specific_calendar_event, **kwargs)
        if action == "get_calendar_view":
            return await run_blocking(client.get_calendar_view, **kwargs)
        if action == "list_calendars":
            return await run_blocking(client.list_calendars, **kwargs)
        if action == "find_meeting_times":
            return await run_blocking(client.find_meeting_times, **kwargs)
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

        resolved = resolve_action(action, _NOTES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_onenote_page_content":
            return await run_blocking(client.get_onenote_page_content, **kwargs)
        if action == "create_onenote_page":
            return await run_blocking(client.create_onenote_page, **kwargs)
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

        resolved = resolve_action(action, _TASKS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_todo_task":
            return await run_blocking(client.get_todo_task, **kwargs)
        if action == "create_todo_task":
            return await run_blocking(client.create_todo_task, **kwargs)
        if action == "update_todo_task":
            return await run_blocking(client.update_todo_task, **kwargs)
        if action == "delete_todo_task":
            return await run_blocking(client.delete_todo_task, **kwargs)
        if action == "get_planner_plan":
            return await run_blocking(client.get_planner_plan, **kwargs)
        if action == "get_planner_task":
            return await run_blocking(client.get_planner_task, **kwargs)
        if action == "create_planner_task":
            return await run_blocking(client.create_planner_task, **kwargs)
        if action == "update_planner_task":
            return await run_blocking(client.update_planner_task, **kwargs)
        if action == "update_planner_task_details":
            return await run_blocking(client.update_planner_task_details, **kwargs)
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

        resolved = resolve_action(action, _CONTACTS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_outlook_contact":
            return await run_blocking(client.get_outlook_contact, **kwargs)
        if action == "create_outlook_contact":
            return await run_blocking(client.create_outlook_contact, **kwargs)
        if action == "update_outlook_contact":
            return await run_blocking(client.update_outlook_contact, **kwargs)
        if action == "delete_outlook_contact":
            return await run_blocking(client.delete_outlook_contact, **kwargs)
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

        resolved = resolve_action(action, _USER_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_current_user":
            return await run_blocking(client.get_current_user, **kwargs)
        if action == "get_me":
            return await run_blocking(client.get_me, **kwargs)
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

        resolved = resolve_action(action, _CHAT_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_chat":
            return await run_blocking(client.get_chat, **kwargs)
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

        resolved = resolve_action(action, _TEAMS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_team":
            return await run_blocking(client.get_team, **kwargs)
        if action == "get_team_channel":
            return await run_blocking(client.get_team_channel, **kwargs)
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

        resolved = resolve_action(action, _SITES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_sites":
            return await run_blocking(client.list_sites, **kwargs)
        if action == "get_site":
            return await run_blocking(client.get_site, **kwargs)
        if action == "get_sharepoint_site_by_path":
            return await run_blocking(client.get_sharepoint_site_by_path, **kwargs)
        if action == "get_sharepoint_sites_delta":
            return await run_blocking(client.get_sharepoint_sites_delta, **kwargs)
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

        resolved = resolve_action(action, _SEARCH_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "search_query":
            return await run_blocking(client.search_query, **kwargs)
        if action == "search_tools":
            return await run_blocking(client.search_tools, **kwargs)
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

        resolved = resolve_action(action, _GROUPS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_groups":
            return await run_blocking(client.list_groups, **kwargs)
        if action == "get_group":
            return await run_blocking(client.get_group, **kwargs)
        if action == "create_group":
            return await run_blocking(client.create_group, **kwargs)
        if action == "update_group":
            return await run_blocking(client.update_group, **kwargs)
        if action == "delete_group":
            return await run_blocking(client.delete_group, **kwargs)
        if action == "list_group_members":
            return await run_blocking(client.list_group_members, **kwargs)
        if action == "add_group_member":
            return await run_blocking(client.add_group_member, **kwargs)
        if action == "remove_group_member":
            return await run_blocking(client.remove_group_member, **kwargs)
        if action == "list_group_owners":
            return await run_blocking(client.list_group_owners, **kwargs)
        if action == "list_group_conversations":
            return await run_blocking(client.list_group_conversations, **kwargs)
        if action == "list_group_drives":
            return await run_blocking(client.list_group_drives, **kwargs)
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

        resolved = resolve_action(action, _ADMIN_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_service_health":
            return await run_blocking(client.list_service_health, **kwargs)
        if action == "get_service_health":
            return await run_blocking(client.get_service_health, **kwargs)
        if action == "list_service_health_issues":
            return await run_blocking(client.list_service_health_issues, **kwargs)
        if action == "get_service_health_issue":
            return await run_blocking(client.get_service_health_issue, **kwargs)
        if action == "list_service_update_messages":
            return await run_blocking(client.list_service_update_messages, **kwargs)
        if action == "get_service_update_message":
            return await run_blocking(client.get_service_update_message, **kwargs)
        if action == "get_admin_sharepoint":
            return await run_blocking(client.get_admin_sharepoint, **kwargs)
        if action == "update_admin_sharepoint":
            return await run_blocking(client.update_admin_sharepoint, **kwargs)
        if action == "list_delegated_admin_relationships":
            return await run_blocking(
                client.list_delegated_admin_relationships, **kwargs
            )
        if action == "get_delegated_admin_relationship":
            return await run_blocking(client.get_delegated_admin_relationship, **kwargs)
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

        resolved = resolve_action(
            action, _ORGANIZATION_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_organization":
            return await run_blocking(client.list_organization, **kwargs)
        if action == "get_organization":
            return await run_blocking(client.get_organization, **kwargs)
        if action == "update_organization":
            return await run_blocking(client.update_organization, **kwargs)
        if action == "get_org_branding":
            return await run_blocking(client.get_org_branding, **kwargs)
        if action == "update_org_branding":
            return await run_blocking(client.update_org_branding, **kwargs)
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

        resolved = resolve_action(action, _DOMAINS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_domains":
            return await run_blocking(client.list_domains, **kwargs)
        if action == "get_domain":
            return await run_blocking(client.get_domain, **kwargs)
        if action == "create_domain":
            return await run_blocking(client.create_domain, **kwargs)
        if action == "delete_domain":
            return await run_blocking(client.delete_domain, **kwargs)
        if action == "verify_domain":
            return await run_blocking(client.verify_domain, **kwargs)
        if action == "list_domain_service_configuration_records":
            return await run_blocking(
                client.list_domain_service_configuration_records, **kwargs
            )
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

        resolved = resolve_action(
            action, _SUBSCRIPTIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_subscriptions":
            return await run_blocking(client.list_subscriptions, **kwargs)
        if action == "get_subscription":
            return await run_blocking(client.get_subscription, **kwargs)
        if action == "create_subscription":
            return await run_blocking(client.create_subscription, **kwargs)
        if action == "update_subscription":
            return await run_blocking(client.update_subscription, **kwargs)
        if action == "delete_subscription":
            return await run_blocking(client.delete_subscription, **kwargs)
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

        resolved = resolve_action(
            action, _COMMUNICATIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_online_meetings":
            return await run_blocking(client.list_online_meetings, **kwargs)
        if action == "get_online_meeting":
            return await run_blocking(client.get_online_meeting, **kwargs)
        if action == "create_online_meeting":
            return await run_blocking(client.create_online_meeting, **kwargs)
        if action == "update_online_meeting":
            return await run_blocking(client.update_online_meeting, **kwargs)
        if action == "delete_online_meeting":
            return await run_blocking(client.delete_online_meeting, **kwargs)
        if action == "list_call_records":
            return await run_blocking(client.list_call_records, **kwargs)
        if action == "get_call_record":
            return await run_blocking(client.get_call_record, **kwargs)
        if action == "list_presences":
            return await run_blocking(client.list_presences, **kwargs)
        if action == "get_presence":
            return await run_blocking(client.get_presence, **kwargs)
        if action == "get_my_presence":
            return await run_blocking(client.get_my_presence, **kwargs)
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

        resolved = resolve_action(action, _IDENTITY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "create_invitation":
            return await run_blocking(client.create_invitation, **kwargs)
        if action == "list_conditional_access_policies":
            return await run_blocking(client.list_conditional_access_policies, **kwargs)
        if action == "get_conditional_access_policy":
            return await run_blocking(client.get_conditional_access_policy, **kwargs)
        if action == "create_conditional_access_policy":
            return await run_blocking(client.create_conditional_access_policy, **kwargs)
        if action == "update_conditional_access_policy":
            return await run_blocking(client.update_conditional_access_policy, **kwargs)
        if action == "delete_conditional_access_policy":
            return await run_blocking(client.delete_conditional_access_policy, **kwargs)
        if action == "list_access_reviews":
            return await run_blocking(client.list_access_reviews, **kwargs)
        if action == "get_access_review":
            return await run_blocking(client.get_access_review, **kwargs)
        if action == "list_entitlement_access_packages":
            return await run_blocking(client.list_entitlement_access_packages, **kwargs)
        if action == "list_lifecycle_workflows":
            return await run_blocking(client.list_lifecycle_workflows, **kwargs)
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

        resolved = resolve_action(action, _SECURITY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_security_alerts":
            return await run_blocking(client.list_security_alerts, **kwargs)
        if action == "get_security_alert":
            return await run_blocking(client.get_security_alert, **kwargs)
        if action == "update_security_alert":
            return await run_blocking(client.update_security_alert, **kwargs)
        if action == "list_security_incidents":
            return await run_blocking(client.list_security_incidents, **kwargs)
        if action == "get_security_incident":
            return await run_blocking(client.get_security_incident, **kwargs)
        if action == "update_security_incident":
            return await run_blocking(client.update_security_incident, **kwargs)
        if action == "list_secure_scores":
            return await run_blocking(client.list_secure_scores, **kwargs)
        if action == "list_threat_intelligence_hosts":
            return await run_blocking(client.list_threat_intelligence_hosts, **kwargs)
        if action == "get_threat_intelligence_host":
            return await run_blocking(client.get_threat_intelligence_host, **kwargs)
        if action == "run_hunting_query":
            return await run_blocking(client.run_hunting_query, **kwargs)
        if action == "list_risk_detections":
            return await run_blocking(client.list_risk_detections, **kwargs)
        if action == "get_risk_detection":
            return await run_blocking(client.get_risk_detection, **kwargs)
        if action == "list_risky_users":
            return await run_blocking(client.list_risky_users, **kwargs)
        if action == "get_risky_user":
            return await run_blocking(client.get_risky_user, **kwargs)
        if action == "dismiss_risky_user":
            return await run_blocking(client.dismiss_risky_user, **kwargs)
        if action == "list_sensitivity_labels":
            return await run_blocking(client.list_sensitivity_labels, **kwargs)
        if action == "get_sensitivity_label":
            return await run_blocking(client.get_sensitivity_label, **kwargs)
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

        resolved = resolve_action(action, _AUDIT_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_directory_audits":
            return await run_blocking(client.list_directory_audits, **kwargs)
        if action == "get_directory_audit":
            return await run_blocking(client.get_directory_audit, **kwargs)
        if action == "list_sign_in_logs":
            return await run_blocking(client.list_sign_in_logs, **kwargs)
        if action == "get_sign_in_log":
            return await run_blocking(client.get_sign_in_log, **kwargs)
        if action == "list_provisioning_logs":
            return await run_blocking(client.list_provisioning_logs, **kwargs)
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

        resolved = resolve_action(action, _REPORTS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_email_activity_report":
            return await run_blocking(client.get_email_activity_report, **kwargs)
        if action == "get_mailbox_usage_report":
            return await run_blocking(client.get_mailbox_usage_report, **kwargs)
        if action == "get_office365_active_users":
            return await run_blocking(client.get_office365_active_users, **kwargs)
        if action == "get_sharepoint_activity_report":
            return await run_blocking(client.get_sharepoint_activity_report, **kwargs)
        if action == "get_teams_user_activity":
            return await run_blocking(client.get_teams_user_activity, **kwargs)
        if action == "get_onedrive_usage_report":
            return await run_blocking(client.get_onedrive_usage_report, **kwargs)
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

        resolved = resolve_action(
            action, _APPLICATIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_applications":
            return await run_blocking(client.list_applications, **kwargs)
        if action == "get_application":
            return await run_blocking(client.get_application, **kwargs)
        if action == "create_application":
            return await run_blocking(client.create_application, **kwargs)
        if action == "update_application":
            return await run_blocking(client.update_application, **kwargs)
        if action == "delete_application":
            return await run_blocking(client.delete_application, **kwargs)
        if action == "add_application_password":
            return await run_blocking(client.add_application_password, **kwargs)
        if action == "remove_application_password":
            return await run_blocking(client.remove_application_password, **kwargs)
        if action == "list_service_principals":
            return await run_blocking(client.list_service_principals, **kwargs)
        if action == "get_service_principal":
            return await run_blocking(client.get_service_principal, **kwargs)
        if action == "create_service_principal":
            return await run_blocking(client.create_service_principal, **kwargs)
        if action == "update_service_principal":
            return await run_blocking(client.update_service_principal, **kwargs)
        if action == "delete_service_principal":
            return await run_blocking(client.delete_service_principal, **kwargs)
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

        resolved = resolve_action(action, _DIRECTORY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_directory_objects":
            return await run_blocking(client.list_directory_objects, **kwargs)
        if action == "get_directory_object":
            return await run_blocking(client.get_directory_object, **kwargs)
        if action == "list_directory_roles":
            return await run_blocking(client.list_directory_roles, **kwargs)
        if action == "get_directory_role":
            return await run_blocking(client.get_directory_role, **kwargs)
        if action == "list_directory_role_templates":
            return await run_blocking(client.list_directory_role_templates, **kwargs)
        if action == "list_deleted_items":
            return await run_blocking(client.list_deleted_items, **kwargs)
        if action == "restore_deleted_item":
            return await run_blocking(client.restore_deleted_item, **kwargs)
        if action == "list_role_definitions":
            return await run_blocking(client.list_role_definitions, **kwargs)
        if action == "get_role_definition":
            return await run_blocking(client.get_role_definition, **kwargs)
        if action == "list_role_assignments":
            return await run_blocking(client.list_role_assignments, **kwargs)
        if action == "get_role_assignment":
            return await run_blocking(client.get_role_assignment, **kwargs)
        if action == "create_role_assignment":
            return await run_blocking(client.create_role_assignment, **kwargs)
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

        resolved = resolve_action(action, _POLICIES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "get_authorization_policy":
            return await run_blocking(client.get_authorization_policy, **kwargs)
        if action == "list_token_lifetime_policies":
            return await run_blocking(client.list_token_lifetime_policies, **kwargs)
        if action == "list_token_issuance_policies":
            return await run_blocking(client.list_token_issuance_policies, **kwargs)
        if action == "list_permission_grant_policies":
            return await run_blocking(client.list_permission_grant_policies, **kwargs)
        if action == "get_admin_consent_policy":
            return await run_blocking(client.get_admin_consent_policy, **kwargs)
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

        resolved = resolve_action(action, _DEVICES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_devices":
            return await run_blocking(client.list_devices, **kwargs)
        if action == "get_device":
            return await run_blocking(client.get_device, **kwargs)
        if action == "delete_device":
            return await run_blocking(client.delete_device, **kwargs)
        if action == "list_managed_devices":
            return await run_blocking(client.list_managed_devices, **kwargs)
        if action == "get_managed_device":
            return await run_blocking(client.get_managed_device, **kwargs)
        if action == "list_device_compliance_policies":
            return await run_blocking(client.list_device_compliance_policies, **kwargs)
        if action == "list_device_configurations":
            return await run_blocking(client.list_device_configurations, **kwargs)
        if action == "wipe_managed_device":
            return await run_blocking(client.wipe_managed_device, **kwargs)
        if action == "retire_managed_device":
            return await run_blocking(client.retire_managed_device, **kwargs)
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

        resolved = resolve_action(action, _EDUCATION_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_education_classes":
            return await run_blocking(client.list_education_classes, **kwargs)
        if action == "get_education_class":
            return await run_blocking(client.get_education_class, **kwargs)
        if action == "list_education_schools":
            return await run_blocking(client.list_education_schools, **kwargs)
        if action == "get_education_school":
            return await run_blocking(client.get_education_school, **kwargs)
        if action == "list_education_users":
            return await run_blocking(client.list_education_users, **kwargs)
        if action == "list_education_assignments":
            return await run_blocking(client.list_education_assignments, **kwargs)
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

        resolved = resolve_action(
            action, _AGREEMENTS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_agreements":
            return await run_blocking(client.list_agreements, **kwargs)
        if action == "get_agreement":
            return await run_blocking(client.get_agreement, **kwargs)
        if action == "create_agreement":
            return await run_blocking(client.create_agreement, **kwargs)
        if action == "delete_agreement":
            return await run_blocking(client.delete_agreement, **kwargs)
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

        resolved = resolve_action(action, _PLACES_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_rooms":
            return await run_blocking(client.list_rooms, **kwargs)
        if action == "list_room_lists":
            return await run_blocking(client.list_room_lists, **kwargs)
        if action == "get_place":
            return await run_blocking(client.get_place, **kwargs)
        if action == "update_place":
            return await run_blocking(client.update_place, **kwargs)
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

        resolved = resolve_action(action, _PRINT_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_printers":
            return await run_blocking(client.list_printers, **kwargs)
        if action == "get_printer":
            return await run_blocking(client.get_printer, **kwargs)
        if action == "list_print_jobs":
            return await run_blocking(client.list_print_jobs, **kwargs)
        if action == "create_print_job":
            return await run_blocking(client.create_print_job, **kwargs)
        if action == "list_print_shares":
            return await run_blocking(client.list_print_shares, **kwargs)
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

        resolved = resolve_action(action, _PRIVACY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_subject_rights_requests":
            return await run_blocking(client.list_subject_rights_requests, **kwargs)
        if action == "get_subject_rights_request":
            return await run_blocking(client.get_subject_rights_request, **kwargs)
        if action == "create_subject_rights_request":
            return await run_blocking(client.create_subject_rights_request, **kwargs)
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

        resolved = resolve_action(action, _SOLUTIONS_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_booking_businesses":
            return await run_blocking(client.list_booking_businesses, **kwargs)
        if action == "get_booking_business":
            return await run_blocking(client.get_booking_business, **kwargs)
        if action == "list_booking_appointments":
            return await run_blocking(client.list_booking_appointments, **kwargs)
        if action == "create_booking_appointment":
            return await run_blocking(client.create_booking_appointment, **kwargs)
        if action == "list_virtual_events":
            return await run_blocking(client.list_virtual_events, **kwargs)
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

        resolved = resolve_action(action, _STORAGE_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_file_storage_containers":
            return await run_blocking(client.list_file_storage_containers, **kwargs)
        if action == "get_file_storage_container":
            return await run_blocking(client.get_file_storage_container, **kwargs)
        if action == "create_file_storage_container":
            return await run_blocking(client.create_file_storage_container, **kwargs)
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

        resolved = resolve_action(
            action, _EMPLOYEE_EXPERIENCE_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_learning_providers":
            return await run_blocking(client.list_learning_providers, **kwargs)
        if action == "get_learning_provider":
            return await run_blocking(client.get_learning_provider, **kwargs)
        if action == "list_learning_course_activities":
            return await run_blocking(client.list_learning_course_activities, **kwargs)
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

        resolved = resolve_action(
            action, _CONNECTIONS_ACTIONS, service="microsoft-agent"
        )
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        if action == "list_external_connections":
            return await run_blocking(client.list_external_connections, **kwargs)
        if action == "get_external_connection":
            return await run_blocking(client.get_external_connection, **kwargs)
        if action == "create_external_connection":
            return await run_blocking(client.create_external_connection, **kwargs)
        if action == "delete_external_connection":
            return await run_blocking(client.delete_external_connection, **kwargs)
        raise ValueError(f"Unknown action: {action}")


def get_mcp_instance() -> tuple[Any, ...]:
    """Initialize and return the MCP instance."""
    load_config()
    args, mcp, middlewares = create_mcp_server(
        name="microsoft-agent MCP",
        version=__version__,
        instructions="microsoft-agent MCP Server — Condensed Action-Routed Tools.",
    )

    @mcp.custom_route("/health", methods=["GET"])
    async def health_check(request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    register_tool_surface(
        mcp,
        client_cls=MicrosoftGraphApi,
        get_client=get_client,
        service="microsoft-agent",
        tools_module=sys.modules[__name__],
    )

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
