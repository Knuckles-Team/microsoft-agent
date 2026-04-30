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

# General urllib3/chardet mismatch warnings
warnings.filterwarnings("ignore", message=".*urllib3.*or chardet.*")
warnings.filterwarnings("ignore", message=".*urllib3.*or charset_normalizer.*")

import logging
import os
import sys
from typing import Any

from agent_utilities.base_utilities import to_boolean
from agent_utilities.mcp_utilities import (
    create_mcp_server,
    ctx_confirm_destructive,
    ctx_progress,
    ctx_sample,
    ctx_set_state,
)
from dotenv import find_dotenv, load_dotenv
from fastmcp import Context, FastMCP
from fastmcp.utilities.logging import get_logger
from pydantic import Field
from starlette.requests import Request
from starlette.responses import JSONResponse

from microsoft_agent.auth import AuthManager, get_client

__version__ = "0.2.57"
print(f"Microsoft MCP v{__version__}")

logger = get_logger(name="TokenMiddleware")
logger.setLevel(logging.DEBUG)


def register_prompts(mcp: FastMCP):
    @mcp.prompt(name="check_email", description="Check your latest emails.")
    def check_email() -> str:
        """Check emails."""
        return "Please check my latest emails."

    @mcp.prompt(
        name="summarize_email", description="Summarize a specific email thread."
    )
    def summarize_email(subject: str) -> str:
        """Summarize email."""
        return f"Please summarize the email thread with subject '{subject}'"

    @mcp.tool(name="calendar_today", description="Show today's calendar events.")
    def calendar_today(
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> str:
        """Show calendar."""
        return "Please show my calendar events for today."


def register_misc_tools(mcp: FastMCP):
    async def health_check(_request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")

    AUTHORITY = "https://login.microsoftonline.com/common"

    SCOPES = [
        "User.Read",
        "Mail.ReadWrite",
        "Calendars.ReadWrite",
        "Files.ReadWrite",
        "Tasks.ReadWrite",
        "Contacts.ReadWrite",
        "Group.ReadWrite.All",
        "Directory.Read.All",
        "Sites.Read.All",
        "Chat.Read",
        "ChatMessage.Read.All",
        "ChannelMessage.Read.All",
        "ServiceHealth.Read.All",
        "ServiceMessage.Read.All",
        "Domain.ReadWrite.All",
        "Organization.ReadWrite.All",
        "OnlineMeetings.ReadWrite",
        "CallRecords.Read.All",
        "Presence.Read.All",
        "User.Invite.All",
        "SecurityEvents.ReadWrite.All",
        "SecurityIncident.ReadWrite.All",
        "ThreatHunting.Read.All",
        "AuditLog.Read.All",
        "Reports.Read.All",
        "Application.ReadWrite.All",
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess",
        "IdentityRiskEvent.Read.All",
        "IdentityRiskyUser.ReadWrite.All",
        "Directory.ReadWrite.All",
        "RoleManagement.ReadWrite.Directory",
        "EntitlementManagement.Read.All",
        "AccessReview.Read.All",
        "LifecycleWorkflows.Read.All",
        "Device.ReadWrite.All",
        "DeviceManagementManagedDevices.ReadWrite.All",
        "DeviceManagementConfiguration.Read.All",
        "EduAssignments.Read",
        "EduRoster.Read",
        "Agreement.ReadWrite.All",
        "Place.Read.All",
        "PrintJob.ReadWriteBasic",
        "Printer.Read.All",
        "SubjectRightsRequest.ReadWrite.All",
        "Bookings.ReadWrite.All",
        "FileStorageContainer.Selected",
        "LearningProvider.Read",
        "ExternalConnection.ReadWrite.All",
        "InformationProtectionPolicy.Read",
        "DelegatedAdminRelationship.Read.All",
    ]

    _ = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)


def register_auth_tools(mcp: FastMCP):
    @mcp.tool(
        name="login",
        description="Authenticate with Microsoft using device code flow",
        tags={"auth"},
    )
    async def login(
        force: bool = Field(
            False, description="Force a new login even if already logged in"
        ),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """Authenticate with Microsoft using device code flow"""
        client = await get_client()
        result = client.login(force=force)
        await ctx_set_state(ctx, "microsoft", "auth_status", "authenticated")
        return result

    @mcp.tool(
        name="logout", description="Log out from Microsoft account", tags={"auth"}
    )
    async def logout(
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """Log out from Microsoft account"""
        if not await ctx_confirm_destructive(ctx, "logout"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return client.logout()

    @mcp.tool(
        name="verify_login",
        description="Check current Microsoft authentication status",
        tags={"auth"},
    )
    async def verify_login(
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """Check current Microsoft authentication status"""
        client = await get_client()
        result = client.verify_login()
        await ctx_set_state(ctx, "microsoft", "auth_token", result.get("jwt") if isinstance(result, dict) else None)
        return result

    @mcp.tool(
        name="list_accounts",
        description="List all available Microsoft accounts",
        tags={"auth"},
    )
    async def list_accounts(
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """List all available Microsoft accounts"""
        client = await get_client()
        return client.list_accounts()


def register_meta_tools(mcp: FastMCP):
    @mcp.tool(
        name="search_tools",
        description="Search available Microsoft Graph API tools",
        tags={"meta"},
    )
    async def search_tools(
        query: str = Field(..., description="Search query"),
        limit: int = Field(20, description="Max results"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """Search available Microsoft Graph API tools"""
        client = await get_client()
        return client.search_tools(query=query, limit=limit)


def register_mail_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_mail_messages",
        description="""list_mail_messages: GET /me/messages

TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:john AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter""",
        tags={"mail", "files", "user"},
    )
    async def list_mail_messages(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_mail_messages: GET /me/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:john AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = await get_client()
        result = await client.list_mail_messages(params=params)
        summary = await ctx_sample(ctx, f"Summarize these emails concisely: {result}")
        if summary and isinstance(result, dict):
            result["ai_summary"] = summary
        return result

    @mcp.tool(
        name="list_mail_folders",
        description="list_mail_folders: GET /me/mailFolders",
        tags={"mail", "files"},
    )
    async def list_mail_folders(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_mail_folders: GET /me/mailFolders"""
        client = await get_client()
        return await client.list_mail_folders(params=params)

    @mcp.tool(
        name="list_mail_folder_messages",
        description="""list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages



TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter""",
        tags={"mail", "files", "user"},
    )
    async def list_mail_folder_messages(
        mailFolder_id: str = Field(..., description="Parameter for mailFolder-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = await get_client()
        return await client.list_mail_folder_messages(
            mailFolder_id=mailFolder_id, params=params
        )

    @mcp.tool(
        name="get_mail_message",
        description="get_mail_message: GET /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    async def get_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_mail_message: GET /me/messages/{message-id}"""
        client = await get_client()
        return await client.get_mail_message(message_id=message_id, params=params)

    @mcp.tool(
        name="send_mail",
        description="""send_mail: POST /me/sendMail



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"mail"},
    )
    async def send_mail(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """send_mail: POST /me/sendMail

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.send_mail(data=data, params=params)

    @mcp.tool(
        name="list_shared_mailbox_messages",
        description="""list_shared_mailbox_messages: GET /users/{user-id}/messages



TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter""",
        tags={"mail", "files", "user"},
    )
    async def list_shared_mailbox_messages(
        user_id: str = Field(..., description="Parameter for user-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_shared_mailbox_messages: GET /users/{user-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = await get_client()
        return await client.list_shared_mailbox_messages(user_id=user_id, params=params)

    @mcp.tool(
        name="list_shared_mailbox_folder_messages",
        description="""list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages



TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter""",
        tags={"mail", "files", "user"},
    )
    async def list_shared_mailbox_folder_messages(
        user_id: str = Field(..., description="Parameter for user-id"),
        mailFolder_id: str = Field(..., description="Parameter for mailFolder-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = await get_client()
        return await client.list_shared_mailbox_folder_messages(
            user_id=user_id, mailFolder_id=mailFolder_id, params=params
        )

    @mcp.tool(
        name="get_shared_mailbox_message",
        description="get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id}",
        tags={"mail", "user"},
    )
    async def get_shared_mailbox_message(
        user_id: str = Field(..., description="Parameter for user-id"),
        message_id: str = Field(..., description="Parameter for message-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id}"""
        client = await get_client()
        return await client.get_shared_mailbox_message(
            user_id=user_id, message_id=message_id, params=params
        )

    @mcp.tool(
        name="send_shared_mailbox_mail",
        description="""send_shared_mailbox_mail: POST /users/{user-id}/sendMail



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"mail"},
    )
    async def send_shared_mailbox_mail(
        user_id: str = Field(..., description="Parameter for user-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """send_shared_mailbox_mail: POST /users/{user-id}/sendMail

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.send_shared_mailbox_mail(
            user_id=user_id, data=data, params=params
        )

    @mcp.tool(
        name="create_draft_email",
        description="create_draft_email: POST /me/messages",
        tags={"mail"},
    )
    async def create_draft_email(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_draft_email: POST /me/messages"""
        client = await get_client()
        return await client.create_draft_email(data=data, params=params)

    @mcp.tool(
        name="delete_mail_message",
        description="delete_mail_message: DELETE /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    async def delete_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_mail_message: DELETE /me/messages/{message-id}"""
        if not await ctx_confirm_destructive(ctx, "delete mail message"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_mail_message(message_id=message_id, params=params)

    @mcp.tool(
        name="move_mail_message",
        description="move_mail_message: POST /me/messages/{message-id}/move",
        tags={"mail", "user"},
    )
    async def move_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """move_mail_message: POST /me/messages/{message-id}/move"""
        client = await get_client()
        return await client.move_mail_message(
            message_id=message_id, data=data, params=params
        )

    @mcp.tool(
        name="update_mail_message",
        description="update_mail_message: PATCH /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    async def update_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_mail_message: PATCH /me/messages/{message-id}"""
        client = await get_client()
        return await client.update_mail_message(
            message_id=message_id, data=data, params=params
        )

    @mcp.tool(
        name="add_mail_attachment",
        description="add_mail_attachment: POST /me/messages/{message-id}/attachments",
        tags={"mail", "user"},
    )
    async def add_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """add_mail_attachment: POST /me/messages/{message-id}/attachments"""
        client = await get_client()
        return await client.add_mail_attachment(
            message_id=message_id, data=data, params=params
        )

    @mcp.tool(
        name="list_mail_attachments",
        description="list_mail_attachments: GET /me/messages/{message-id}/attachments",
        tags={"mail", "files", "user"},
    )
    async def list_mail_attachments(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_mail_attachments: GET /me/messages/{message-id}/attachments"""
        client = await get_client()
        return await client.list_mail_attachments(message_id=message_id, params=params)

    @mcp.tool(
        name="get_mail_attachment",
        description="get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id}",
        tags={"mail", "user"},
    )
    async def get_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        attachment_id: str = Field(..., description="Parameter for attachment-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id}"""
        client = await get_client()
        return await client.get_mail_attachment(
            message_id=message_id, attachment_id=attachment_id, params=params
        )

    @mcp.tool(
        name="delete_mail_attachment",
        description="delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id}",
        tags={"mail", "user"},
    )
    async def delete_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        attachment_id: str = Field(..., description="Parameter for attachment-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id}"""
        if not await ctx_confirm_destructive(ctx, "delete mail attachment"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_mail_attachment(
            message_id=message_id, attachment_id=attachment_id, params=params
        )

    @mcp.tool(
        name="get_root_folder",
        description="get_root_folder: GET /drives/{drive-id}/root",
        tags={"mail", "files"},
    )
    async def get_root_folder(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_root_folder: GET /drives/{drive-id}/root"""
        client = await get_client()
        return await client.get_root_folder(drive_id=drive_id, params=params)

    @mcp.tool(
        name="list_folder_files",
        description="list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children",
        tags={"mail", "files"},
    )
    async def list_folder_files(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children"""
        client = await get_client()
        return await client.list_folder_files(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="list_chat_messages",
        description="list_chat_messages: GET /chats/{chat-id}/messages",
        tags={"mail", "files", "user", "chat"},
    )
    async def list_chat_messages(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_chat_messages: GET /chats/{chat-id}/messages"""
        client = await get_client()
        return await client.list_chat_messages(chat_id=chat_id, params=params)

    @mcp.tool(
        name="get_chat_message",
        description="get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id}",
        tags={"mail", "user", "chat"},
    )
    async def get_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id}"""
        client = await get_client()
        return await client.get_chat_message(
            chat_id=chat_id, chatMessage_id=chatMessage_id, params=params
        )

    @mcp.tool(
        name="send_chat_message",
        description="send_chat_message: POST /chats/{chat-id}/messages",
        tags={"mail", "user", "chat"},
    )
    async def send_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """send_chat_message: POST /chats/{chat-id}/messages"""
        client = await get_client()
        return await client.send_chat_message(chat_id=chat_id, data=data, params=params)

    @mcp.tool(
        name="list_channel_messages",
        description="list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages",
        tags={"mail", "files", "user", "teams"},
    )
    async def list_channel_messages(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages"""
        client = await get_client()
        return await client.list_channel_messages(
            team_id=team_id, channel_id=channel_id, params=params
        )

    @mcp.tool(
        name="get_channel_message",
        description="get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}",
        tags={"mail", "user", "teams"},
    )
    async def get_channel_message(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}"""
        client = await get_client()
        return await client.get_channel_message(
            team_id=team_id,
            channel_id=channel_id,
            chatMessage_id=chatMessage_id,
            params=params,
        )

    @mcp.tool(
        name="send_channel_message",
        description="send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages",
        tags={"mail", "user", "teams"},
    )
    async def send_channel_message(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages"""
        client = await get_client()
        return await client.send_channel_message(
            team_id=team_id, channel_id=channel_id, data=data, params=params
        )

    @mcp.tool(
        name="list_chat_message_replies",
        description="list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies",
        tags={"mail", "files", "user", "chat"},
    )
    async def list_chat_message_replies(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies"""
        client = await get_client()
        return await client.list_chat_message_replies(
            chat_id=chat_id, chatMessage_id=chatMessage_id, params=params
        )

    @mcp.tool(
        name="reply_to_chat_message",
        description="reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies",
        tags={"mail", "user", "chat"},
    )
    async def reply_to_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies"""
        client = await get_client()
        return await client.reply_to_chat_message(
            chat_id=chat_id, chatMessage_id=chatMessage_id, data=data, params=params
        )


def register_files_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_users",
        description="""list_users: GET /users



TIP: CRITICAL: This request requires the ConsistencyLevel header set to eventual. When searching users, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'displayName:'. Examples: $search='displayName:john' | $search='displayName:john' OR 'displayName:jane'. Remember: ALWAYS wrap the entire search expression in double quotes and set the ConsistencyLevel header to eventual! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter""",
        tags={"files", "user"},
    )
    async def list_users(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_users: GET /users

        TIP: CRITICAL: This request requires the ConsistencyLevel header set to eventual. When searching users, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'displayName:'. Examples: $search='displayName:john' | $search='displayName:john' OR 'displayName:jane'. Remember: ALWAYS wrap the entire search expression in double quotes and set the ConsistencyLevel header to eventual! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return await client.list_users(params=params)

    @mcp.tool(
        name="list_drives", description="list_drives: GET /me/drives", tags={"files"}
    )
    async def list_drives(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_drives: GET /me/drives"""
        client = await get_client()
        return await client.list_drives(params=params)

    @mcp.tool(
        name="get_drive_root_item",
        description="get_drive_root_item: GET /drives/{drive-id}/root",
        tags={"files"},
    )
    async def get_drive_root_item(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_drive_root_item: GET /drives/{drive-id}/root"""
        client = await get_client()
        return await client.get_drive_root_item(drive_id=drive_id, params=params)

    @mcp.tool(
        name="download_onedrive_file_content",
        description="download_onedrive_file_content: GET /drives/{drive-id}/items/{driveItem-id}/content",
        tags={"files"},
    )
    async def download_onedrive_file_content(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        await ctx_progress(ctx, 0, 100)
        """download_onedrive_file_content: GET /drives/{drive-id}/items/{driveItem-id}/content"""
        client = await get_client()
        await ctx_progress(ctx, 100, 100)
        return await client.download_onedrive_file_content(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="delete_onedrive_file",
        description="delete_onedrive_file: DELETE /drives/{drive-id}/items/{driveItem-id}",
        tags={"files"},
    )
    async def delete_onedrive_file(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_onedrive_file: DELETE /drives/{drive-id}/items/{driveItem-id}"""
        if not await ctx_confirm_destructive(ctx, "delete onedrive file"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_onedrive_file(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="upload_file_content",
        description="upload_file_content: PUT /drives/{drive-id}/items/{driveItem-id}/content",
        tags={"files"},
    )
    async def upload_file_content(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        await ctx_progress(ctx, 0, 100)
        """upload_file_content: PUT /drives/{drive-id}/items/{driveItem-id}/content"""
        client = await get_client()
        await ctx_progress(ctx, 100, 100)
        return await client.upload_file_content(
            drive_id=drive_id, driveItem_id=driveItem_id, data=data, params=params
        )

    @mcp.tool(
        name="create_excel_chart",
        description="create_excel_chart: POST /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add",
        tags={"files"},
    )
    async def create_excel_chart(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_excel_chart: POST /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add"""
        client = await get_client()
        return await client.create_excel_chart(
            drive_id=drive_id,
            item_id=driveItem_id,
            worksheet_id=workbookWorksheet_id,
            data=data,
            params=params,
        )

    @mcp.tool(
        name="format_excel_range",
        description="format_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/format",
        tags={"files"},
    )
    async def format_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """format_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/format"""
        client = await get_client()
        return await client.format_excel_range(
            drive_id=drive_id,
            worksheet_id=workbookWorksheet_id,
            item_id=driveItem_id,
            address="",
            data=data,
            params=params,
        )

    @mcp.tool(
        name="sort_excel_range",
        description="sort_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/sort",
        tags={"files"},
    )
    async def sort_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """sort_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/sort"""
        client = await get_client()
        return await client.sort_excel_range(
            drive_id=drive_id,
            item_id=driveItem_id,
            worksheet_id=workbookWorksheet_id,
            address="",
            data=data,
            params=params,
        )

    @mcp.tool(
        name="get_excel_range",
        description="get_excel_range: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
        tags={"files"},
    )
    async def get_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        address: str = Field(..., description="Parameter for address"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_excel_range: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')"""
        client = await get_client()
        return await client.get_excel_range(
            drive_id=drive_id,
            item_id=driveItem_id,
            worksheet_id=workbookWorksheet_id,
            address=address,
            params=params,
        )

    @mcp.tool(
        name="list_excel_worksheets",
        description="list_excel_worksheets: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets",
        tags={"files"},
    )
    async def list_excel_worksheets(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_excel_worksheets: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets"""
        client = await get_client()
        return await client.list_excel_worksheets(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="list_excel_tables",
        description="""list_excel_tables: GET /drives/{drive-id}/items/{driveItem-id}/workbook/tables



List Excel tables in a workbook.""",
        tags={"files"},
    )
    async def list_excel_tables(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        item_id: str = Field(..., description="Parameter for driveItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_excel_tables: GET /drives/{drive-id}/items/{driveItem-id}/workbook/tables"""
        client = await get_client()
        return await client.list_excel_tables(
            drive_id=drive_id, item_id=item_id, params=params
        )

    @mcp.tool(
        name="get_excel_workbook",
        description="get_excel_workbook: GET /drives/{drive-id}/items/{item-id}/workbook",
        tags={"files"},
    )
    async def get_excel_workbook(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        item_id: str = Field(..., description="Parameter for item-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_excel_workbook: GET /drives/{drive-id}/items/{item-id}/workbook"""
        client = await get_client()
        return await client.get_excel_workbook(
            drive_id=drive_id, item_id=item_id, params=params
        )

    @mcp.tool(
        name="list_onenote_notebooks",
        description="list_onenote_notebooks: GET /me/onenote/notebooks",
        tags={"files", "notes"},
    )
    async def list_onenote_notebooks(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_onenote_notebooks: GET /me/onenote/notebooks"""
        client = await get_client()
        return await client.list_onenote_notebooks(params=params)

    @mcp.tool(
        name="list_onenote_notebook_sections",
        description="list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections",
        tags={"files", "notes"},
    )
    async def list_onenote_notebook_sections(
        notebook_id: str = Field(..., description="Parameter for notebook-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections"""
        client = await get_client()
        return await client.list_onenote_notebook_sections(
            notebook_id=notebook_id, params=params
        )

    @mcp.tool(
        name="list_onenote_section_pages",
        description="list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages",
        tags={"files", "notes"},
    )
    async def list_onenote_section_pages(
        onenoteSection_id: str = Field(
            ..., description="Parameter for onenoteSection-id"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages"""
        client = await get_client()
        return await client.list_onenote_section_pages(
            onenoteSection_id=onenoteSection_id, params=params
        )

    @mcp.tool(
        name="list_todo_task_lists",
        description="list_todo_task_lists: GET /me/todo/lists",
        tags={"files", "tasks"},
    )
    async def list_todo_task_lists(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_todo_task_lists: GET /me/todo/lists"""
        client = await get_client()
        return await client.list_todo_task_lists(params=params)

    @mcp.tool(
        name="list_todo_tasks",
        description="list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks",
        tags={"files", "tasks"},
    )
    async def list_todo_tasks(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks"""
        client = await get_client()
        return await client.list_todo_tasks(
            todoTaskList_id=todoTaskList_id, params=params
        )

    @mcp.tool(
        name="list_planner_tasks",
        description="list_planner_tasks: GET /me/planner/tasks",
        tags={"files", "tasks"},
    )
    async def list_planner_tasks(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_planner_tasks: GET /me/planner/tasks"""
        client = await get_client()
        return await client.list_planner_tasks(params=params)

    @mcp.tool(
        name="list_plan_tasks",
        description="list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks",
        tags={"files", "tasks"},
    )
    async def list_plan_tasks(
        plannerPlan_id: str = Field(..., description="Parameter for plannerPlan-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks"""
        client = await get_client()
        return await client.list_plan_tasks(
            plannerPlan_id=plannerPlan_id, params=params
        )

    @mcp.tool(
        name="list_outlook_contacts",
        description="list_outlook_contacts: GET /me/contacts",
        tags={"files", "contacts"},
    )
    async def list_outlook_contacts(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_outlook_contacts: GET /me/contacts"""
        client = await get_client()
        return await client.list_outlook_contacts(params=params)

    @mcp.tool(
        name="list_chats",
        description="list_chats: GET /me/chats",
        tags={"files", "chat"},
    )
    async def list_chats(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_chats: GET /me/chats"""
        client = await get_client()
        return await client.list_chats(params=params)

    @mcp.tool(
        name="get_excel_worksheet",
        description="get_excel_worksheet: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}",
        tags={"files"},
    )
    async def get_excel_worksheet(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        item_id: str = Field(..., description="Parameter for item-id"),
        worksheet_id: str = Field(..., description="Parameter for worksheet-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_excel_worksheet: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id}"""
        client = await get_client()
        return await client.get_excel_worksheet(
            drive_id=drive_id,
            item_id=item_id,
            worksheet_id=worksheet_id,
            params=params,
        )

    @mcp.tool(
        name="list_joined_teams",
        description="list_joined_teams: GET /me/joinedTeams",
        tags={"files", "teams"},
    )
    async def list_joined_teams(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_joined_teams: GET /me/joinedTeams"""
        client = await get_client()
        return await client.list_joined_teams(params=params)

    @mcp.tool(
        name="list_team_channels",
        description="list_team_channels: GET /teams/{team-id}/channels",
        tags={"files", "teams"},
    )
    async def list_team_channels(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_team_channels: GET /teams/{team-id}/channels"""
        client = await get_client()
        return await client.list_team_channels(team_id=team_id, params=params)

    @mcp.tool(
        name="list_team_members",
        description="list_team_members: GET /teams/{team-id}/members",
        tags={"files", "user", "teams"},
    )
    async def list_team_members(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_team_members: GET /teams/{team-id}/members"""
        client = await get_client()
        return await client.list_team_members(team_id=team_id, params=params)

    @mcp.tool(
        name="list_site_drives",
        description="list_site_drives: GET /sites/{site-id}/drives",
        tags={"files", "sites"},
    )
    async def list_site_drives(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_site_drives: GET /sites/{site-id}/drives"""
        client = await get_client()
        return await client.list_site_drives(site_id=site_id, params=params)

    @mcp.tool(
        name="get_site_drive_by_id",
        description="get_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id}",
        tags={"files", "sites"},
    )
    async def get_site_drive_by_id(
        site_id: str = Field(..., description="Parameter for site-id"),
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id}"""
        client = await get_client()
        return await client.get_site_drive_by_id(
            site_id=site_id, drive_id=drive_id, params=params
        )

    @mcp.tool(
        name="list_site_items",
        description="list_site_items: GET /sites/{site-id}/items",
        tags={"files", "sites"},
    )
    async def list_site_items(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_site_items: GET /sites/{site-id}/items"""
        client = await get_client()
        return await client.list_site_items(site_id=site_id, params=params)

    @mcp.tool(
        name="get_site_item",
        description="get_site_item: GET /sites/{site-id}/items/{baseItem-id}",
        tags={"files", "sites"},
    )
    async def get_site_item(
        site_id: str = Field(..., description="Parameter for site-id"),
        baseItem_id: str = Field(..., description="Parameter for baseItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_site_item: GET /sites/{site-id}/items/{baseItem-id}"""
        client = await get_client()
        return await client.get_site_item(
            site_id=site_id, baseItem_id=baseItem_id, params=params
        )

    @mcp.tool(
        name="list_site_lists",
        description="""list_site_lists: GET /sites/{site-id}/lists



List lists for a SharePoint site.""",
        tags={"files", "sites"},
    )
    async def list_site_lists(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_site_lists: GET /sites/{site-id}/lists"""
        client = await get_client()
        return await client.list_site_lists(site_id=site_id, params=params)

    @mcp.tool(
        name="get_site_list",
        description="""get_site_list: GET /sites/{site-id}/lists/{list-id}



Get a specific SharePoint site list.""",
        tags={"files", "sites"},
    )
    async def get_site_list(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_site_list: GET /sites/{site-id}/lists/{list-id}"""
        client = await get_client()
        return await client.get_site_list(
            site_id=site_id, list_id=list_id, params=params
        )

    @mcp.tool(
        name="list_sharepoint_site_list_items",
        description="""list_sharepoint_site_list_items: GET /sites/{site-id}/lists/{list-id}/items



List items in a SharePoint site list.""",
        tags={"files", "sites"},
    )
    async def list_sharepoint_site_list_items(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_sharepoint_site_list_items: GET /sites/{site-id}/lists/{list-id}/items"""
        client = await get_client()
        return await client.list_sharepoint_site_list_items(
            site_id=site_id, list_id=list_id, params=params
        )

    @mcp.tool(
        name="get_sharepoint_site_list_item",
        description="get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id}",
        tags={"files", "sites"},
    )
    async def get_sharepoint_site_list_item(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        listItem_id: str = Field(..., description="Parameter for listItem-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id}"""
        client = await get_client()
        return await client.get_sharepoint_site_list_item(
            site_id=site_id, list_id=list_id, listItem_id=listItem_id, params=params
        )

    @mcp.tool(
        name="get_excel_table",
        description="get_excel_table: GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}",
        tags={"files"},
    )
    async def get_excel_table(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        item_id: str = Field(..., description="Parameter for item-id"),
        table_id: str = Field(..., description="Parameter for table-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_excel_table: GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id}"""
        client = await get_client()
        return await client.get_excel_table(
            drive_id=drive_id, item_id=item_id, table_id=table_id, params=params
        )


def register_calendar_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_calendar_events",
        description="list_calendar_events: GET /me/events",
        tags={"calendar", "files"},
    )
    async def list_calendar_events(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        timezone: str | None = Field(None, description="IANA timezone"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_calendar_events: GET /me/events"""
        client = await get_client()
        return await client.list_calendar_events(params=params, timezone=timezone)

    @mcp.tool(
        name="get_calendar_event",
        description="get_calendar_event: GET /me/events/{event-id}",
        tags={"calendar"},
    )
    async def get_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        timezone: str | None = Field(None, description="IANA timezone"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_calendar_event: GET /me/events/{event-id}"""
        client = await get_client()
        return await client.get_calendar_event(
            event_id=event_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="create_calendar_event",
        description="""create_calendar_event: POST /me/events



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"calendar"},
    )
    async def create_calendar_event(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_calendar_event: POST /me/events

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.create_calendar_event(data=data, params=params)

    @mcp.tool(
        name="update_calendar_event",
        description="""update_calendar_event: PATCH /me/events/{event-id}



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"calendar"},
    )
    async def update_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_calendar_event: PATCH /me/events/{event-id}

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.update_calendar_event(
            event_id=event_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_calendar_event",
        description="delete_calendar_event: DELETE /me/events/{event-id}",
        tags={"calendar"},
    )
    async def delete_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_calendar_event: DELETE /me/events/{event-id}"""
        if not await ctx_confirm_destructive(ctx, "delete calendar event"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_calendar_event(event_id=event_id, params=params)

    @mcp.tool(
        name="list_specific_calendar_events",
        description="list_specific_calendar_events: GET /me/calendars/{calendar-id}/events",
        tags={"calendar", "files"},
    )
    async def list_specific_calendar_events(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        timezone: str | None = Field(None, description="IANA timezone"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_specific_calendar_events: GET /me/calendars/{calendar-id}/events"""
        client = await get_client()

        return await client.list_specific_calendar_events(
            calendar_id=calendar_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="get_specific_calendar_event",
        description="get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id}",
        tags={"calendar"},
    )
    async def get_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        timezone: str | None = Field(None, description="IANA timezone"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id}"""
        client = await get_client()
        return await client.get_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="create_specific_calendar_event",
        description="""create_specific_calendar_event: POST /me/calendars/{calendar-id}/events



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"calendar"},
    )
    async def create_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_specific_calendar_event: POST /me/calendars/{calendar-id}/events

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.create_specific_calendar_event(
            calendar_id=calendar_id, data=data, params=params
        )

    @mcp.tool(
        name="update_specific_calendar_event",
        description="""update_specific_calendar_event: PATCH /me/calendars/{calendar-id}/events/{event-id}



TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.""",
        tags={"calendar"},
    )
    async def update_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_specific_calendar_event: PATCH /me/calendars/{calendar-id}/events/{event-id}

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = await get_client()
        return await client.update_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_specific_calendar_event",
        description="delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id}",
        tags={"calendar"},
    )
    async def delete_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id}"""
        if not await ctx_confirm_destructive(ctx, "delete specific calendar event"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, params=params
        )

    @mcp.tool(
        name="get_calendar_view",
        description="get_calendar_view: GET /me/calendarView",
        tags={"calendar"},
    )
    async def get_calendar_view(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        timezone: str | None = Field(None, description="IANA timezone"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_calendar_view: GET /me/calendarView"""
        client = await get_client()
        return await client.get_calendar_view(params=params, timezone=timezone)

    @mcp.tool(
        name="list_calendars",
        description="list_calendars: GET /me/calendars",
        tags={"calendar", "files"},
    )
    async def list_calendars(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_calendars: GET /me/calendars"""
        client = await get_client()
        return await client.list_calendars(params=params)

    @mcp.tool(
        name="find_meeting_times",
        description="find_meeting_times: POST /me/findMeetingTimes",
        tags={"calendar", "user"},
    )
    async def find_meeting_times(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """find_meeting_times: POST /me/findMeetingTimes"""
        client = await get_client()
        return await client.find_meeting_times(data=data, params=params)


def register_notes_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_onenote_page_content",
        description="get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content",
        tags={"notes"},
    )
    async def get_onenote_page_content(
        onenotePage_id: str = Field(..., description="Parameter for onenotePage-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content"""
        client = await get_client()
        return await client.get_onenote_page_content(
            onenotePage_id=onenotePage_id, params=params
        )

    @mcp.tool(
        name="create_onenote_page",
        description="create_onenote_page: POST /me/onenote/pages",
        tags={"notes"},
    )
    async def create_onenote_page(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_onenote_page: POST /me/onenote/pages"""
        client = await get_client()
        return await client.create_onenote_page(data=data, params=params)


def register_tasks_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_todo_task",
        description="get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}",
        tags={"tasks"},
    )
    async def get_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        client = await get_client()
        return await client.get_todo_task(
            todoTaskList_id=todoTaskList_id, todoTask_id=todoTask_id, params=params
        )

    @mcp.tool(
        name="create_todo_task",
        description="create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks",
        tags={"tasks"},
    )
    async def create_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks"""
        client = await get_client()
        return await client.create_todo_task(
            todoTaskList_id=todoTaskList_id, data=data, params=params
        )

    @mcp.tool(
        name="update_todo_task",
        description="update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}",
        tags={"tasks"},
    )
    async def update_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        client = await get_client()
        return await client.update_todo_task(
            todoTaskList_id=todoTaskList_id,
            todoTask_id=todoTask_id,
            data=data,
            params=params,
        )

    @mcp.tool(
        name="delete_todo_task",
        description="delete_todo_task: DELETE /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}",
        tags={"tasks"},
    )
    async def delete_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_todo_task: DELETE /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        if not await ctx_confirm_destructive(ctx, "delete todo task"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_todo_task(
            todoTaskList_id=todoTaskList_id, todoTask_id=todoTask_id, params=params
        )

    @mcp.tool(
        name="get_planner_plan",
        description="get_planner_plan: GET /planner/plans/{plannerPlan-id}",
        tags={"tasks"},
    )
    async def get_planner_plan(
        plannerPlan_id: str = Field(..., description="Parameter for plannerPlan-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_planner_plan: GET /planner/plans/{plannerPlan-id}"""
        client = await get_client()
        return await client.get_planner_plan(
            plannerPlan_id=plannerPlan_id, params=params
        )

    @mcp.tool(
        name="get_planner_task",
        description="get_planner_task: GET /planner/tasks/{plannerTask-id}",
        tags={"tasks"},
    )
    async def get_planner_task(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_planner_task: GET /planner/tasks/{plannerTask-id}"""
        client = await get_client()
        return await client.get_planner_task(
            plannerTask_id=plannerTask_id, params=params
        )

    @mcp.tool(
        name="create_planner_task",
        description="create_planner_task: POST /planner/tasks",
        tags={"tasks"},
    )
    async def create_planner_task(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_planner_task: POST /planner/tasks"""
        client = await get_client()
        return await client.create_planner_task(data=data, params=params)

    @mcp.tool(
        name="update_planner_task",
        description="update_planner_task: PATCH /planner/tasks/{plannerTask-id}",
        tags={"tasks"},
    )
    async def update_planner_task(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_planner_task: PATCH /planner/tasks/{plannerTask-id}"""
        client = await get_client()
        return await client.update_planner_task(
            plannerTask_id=plannerTask_id, data=data, params=params
        )

    @mcp.tool(
        name="update_planner_task_details",
        description="update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details",
        tags={"tasks"},
    )
    async def update_planner_task_details(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details"""
        client = await get_client()
        return await client.update_planner_task_details(
            plannerTask_id=plannerTask_id, data=data, params=params
        )


def register_contacts_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_outlook_contact",
        description="get_outlook_contact: GET /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    async def get_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_outlook_contact: GET /me/contacts/{contact-id}"""
        client = await get_client()
        return await client.get_outlook_contact(contact_id=contact_id, params=params)

    @mcp.tool(
        name="create_outlook_contact",
        description="create_outlook_contact: POST /me/contacts",
        tags={"contacts"},
    )
    async def create_outlook_contact(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_outlook_contact: POST /me/contacts"""
        client = await get_client()
        return await client.create_outlook_contact(data=data, params=params)

    @mcp.tool(
        name="update_outlook_contact",
        description="update_outlook_contact: PATCH /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    async def update_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_outlook_contact: PATCH /me/contacts/{contact-id}"""
        client = await get_client()
        return await client.update_outlook_contact(
            contact_id=contact_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_outlook_contact",
        description="delete_outlook_contact: DELETE /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    async def delete_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_outlook_contact: DELETE /me/contacts/{contact-id}"""
        if not await ctx_confirm_destructive(ctx, "delete outlook contact"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_outlook_contact(contact_id=contact_id, params=params)


def register_user_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_current_user", description="get_current_user: GET /me", tags={"user"}
    )
    async def get_current_user(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_current_user: GET /me"""
        client = await get_client()
        return await client.get_current_user(params=params)

    @mcp.tool(
        name="get_me",
        description="get_me: GET /me",
        tags={"user"},
    )
    async def get_me(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_me: GET /me"""
        client = await get_client()
        return await client.get_me(params=params)


def register_chat_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_chat", description="get_chat: GET /chats/{chat-id}", tags={"chat"}
    )
    async def get_chat(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_chat: GET /chats/{chat-id}"""
        client = await get_client()
        return await client.get_chat(chat_id=chat_id, params=params)


def register_teams_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_team", description="get_team: GET /teams/{team-id}", tags={"teams"}
    )
    async def get_team(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_team: GET /teams/{team-id}"""
        client = await get_client()
        return await client.get_team(team_id=team_id, params=params)

    @mcp.tool(
        name="get_team_channel",
        description="get_team_channel: GET /teams/{team-id}/channels/{channel-id}",
        tags={"teams"},
    )
    async def get_team_channel(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_team_channel: GET /teams/{team-id}/channels/{channel-id}"""
        client = await get_client()
        return await client.get_team_channel(
            team_id=team_id, channel_id=channel_id, params=params
        )


def register_sites_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_sites",
        description="list_sites: GET /sites",
        tags={"sites"},
    )
    async def list_sites(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_sites: GET /sites"""
        client = await get_client()
        return await client.list_sites(params=params)

    @mcp.tool(
        name="get_site",
        description="get_site: GET /sites/{site-id}",
        tags={"sites"},
    )
    async def get_site(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_site: GET /sites/{site-id}"""
        client = await get_client()
        return await client.get_site(site_id=site_id, params=params)

    @mcp.tool(
        name="get_sharepoint_site_by_path",
        description="get_sharepoint_site_by_path: GET /sites/{site-id}/getByPath(path='{path}')",
        tags={"sites"},
    )
    async def get_sharepoint_site_by_path(
        site_id: str = Field(..., description="Parameter for site-id"),
        path: str = Field(..., description="Parameter for path"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sharepoint_site_by_path: GET /sites/{site-id}/getByPath(path='{path}')"""
        client = await get_client()
        return await client.get_sharepoint_site_by_path(
            site_id=site_id, path=path, params=params
        )

    @mcp.tool(
        name="get_sharepoint_sites_delta",
        description="get_sharepoint_sites_delta: GET /sites/delta()",
        tags={"sites"},
    )
    async def get_sharepoint_sites_delta(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sharepoint_sites_delta: GET /sites/delta()"""
        client = await get_client()
        return await client.get_sharepoint_sites_delta(params=params)


def register_search_tools(mcp: FastMCP):
    @mcp.tool(
        name="search_query",
        description="search_query: POST /search/query",
        tags={"search"},
    )
    async def search_query(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """search_query: POST /search/query"""
        client = await get_client()
        result = await client.search_query(data=data, params=params)
        summary = await ctx_sample(ctx, f"Summarize these Microsoft search results concisely: {result}")
        if summary:
            if isinstance(result, dict):
                result["ai_summary"] = summary
            elif isinstance(result, list):
                result = {"results": result, "ai_summary": summary}
        return result


def register_groups_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_groups",
        description="""list_groups: GET /groups



List all Microsoft 365 groups and security groups in the organization. Supports $filter, $search, $select, $top, $orderby, $count query parameters. Requires ConsistencyLevel: eventual header for advanced queries.""",
        tags={"groups"},
    )
    async def list_groups(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_groups: GET /groups"""
        client = await get_client()
        return await client.list_groups(params=params)

    @mcp.tool(
        name="get_group",
        description="""get_group: GET /groups/{group-id}



Get properties and relationships of a group object.""",
        tags={"groups"},
    )
    async def get_group(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_group: GET /groups/{group-id}"""
        client = await get_client()
        return await client.get_group(group_id=group_id, params=params)

    @mcp.tool(
        name="create_group",
        description="""create_group: POST /groups



Create a new Microsoft 365 group or security group. Required fields: displayName, mailNickname, mailEnabled, securityEnabled. For M365 groups, set groupTypes=['Unified'].""",
        tags={"groups"},
    )
    async def create_group(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_group: POST /groups"""
        client = await get_client()
        return await client.create_group(data=data, params=params)

    @mcp.tool(
        name="update_group",
        description="""update_group: PATCH /groups/{group-id}



Update properties of a group object.""",
        tags={"groups"},
    )
    async def update_group(
        group_id: str = Field(..., description="Parameter for group-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_group: PATCH /groups/{group-id}"""
        client = await get_client()
        return await client.update_group(group_id=group_id, data=data, params=params)

    @mcp.tool(
        name="delete_group",
        description="""delete_group: DELETE /groups/{group-id}



Delete a group. This permanently removes the group and its associated content.""",
        tags={"groups"},
    )
    async def delete_group(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_group: DELETE /groups/{group-id}"""
        if not await ctx_confirm_destructive(ctx, "delete group"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_group(group_id=group_id, params=params)

    @mcp.tool(
        name="list_group_members",
        description="""list_group_members: GET /groups/{group-id}/members



Get a list of the group's direct members.""",
        tags={"groups", "user"},
    )
    async def list_group_members(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_group_members: GET /groups/{group-id}/members"""
        client = await get_client()
        return await client.list_group_members(group_id=group_id, params=params)

    @mcp.tool(
        name="add_group_member",
        description="""add_group_member: POST /groups/{group-id}/members/$ref



Add a member to a group. Provide memberId or directoryObjectId in the request body.""",
        tags={"groups", "user"},
    )
    async def add_group_member(
        group_id: str = Field(..., description="Parameter for group-id"),
        data: dict[str, Any] | None = Field(
            None, description="Request body data with memberId"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """add_group_member: POST /groups/{group-id}/members/$ref"""
        client = await get_client()
        return await client.add_group_member(
            group_id=group_id, data=data, params=params
        )

    @mcp.tool(
        name="remove_group_member",
        description="""remove_group_member: DELETE /groups/{group-id}/members/{member-id}/$ref



Remove a member from a group.""",
        tags={"groups", "user"},
    )
    async def remove_group_member(
        group_id: str = Field(..., description="Parameter for group-id"),
        member_id: str = Field(
            ..., description="Parameter for member-id (directory object ID)"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """remove_group_member: DELETE /groups/{group-id}/members/{member-id}/$ref"""
        if not await ctx_confirm_destructive(ctx, "remove group member"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.remove_group_member(
            group_id=group_id, member_id=member_id, params=params
        )

    @mcp.tool(
        name="list_group_owners",
        description="""list_group_owners: GET /groups/{group-id}/owners



Get owners of a group.""",
        tags={"groups", "user"},
    )
    async def list_group_owners(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_group_owners: GET /groups/{group-id}/owners"""
        client = await get_client()
        return await client.list_group_owners(group_id=group_id, params=params)

    @mcp.tool(
        name="list_group_conversations",
        description="""list_group_conversations: GET /groups/{group-id}/conversations



List conversations in a Microsoft 365 group.""",
        tags={"groups", "chat"},
    )
    async def list_group_conversations(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_group_conversations: GET /groups/{group-id}/conversations"""
        client = await get_client()
        return await client.list_group_conversations(group_id=group_id, params=params)

    @mcp.tool(
        name="list_group_drives",
        description="""list_group_drives: GET /groups/{group-id}/drives



List drives (document libraries) of a group.""",
        tags={"groups", "files"},
    )
    async def list_group_drives(
        group_id: str = Field(..., description="Parameter for group-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_group_drives: GET /groups/{group-id}/drives"""
        client = await get_client()
        return await client.list_group_drives(group_id=group_id, params=params)


def register_admin_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_service_health",
        description="""list_service_health: GET /admin/serviceAnnouncement/healthOverviews



Get the service health status for all services in the tenant.""",
        tags={"admin"},
    )
    async def list_service_health(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_service_health: GET /admin/serviceAnnouncement/healthOverviews"""
        client = await get_client()
        return await client.list_service_health(params=params)

    @mcp.tool(
        name="get_service_health",
        description="""get_service_health: GET /admin/serviceAnnouncement/healthOverviews/{service-name}



Get the health status for a specific service.""",
        tags={"admin"},
    )
    async def get_service_health(
        service_name: str = Field(
            ..., description="Service name (e.g. 'Exchange Online')"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_service_health: GET /admin/serviceAnnouncement/healthOverviews/{service-name}"""
        client = await get_client()
        return await client.get_service_health(service_name=service_name, params=params)

    @mcp.tool(
        name="list_service_health_issues",
        description="""list_service_health_issues: GET /admin/serviceAnnouncement/issues



List all service health issues for the tenant.""",
        tags={"admin"},
    )
    async def list_service_health_issues(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_service_health_issues: GET /admin/serviceAnnouncement/issues"""
        client = await get_client()
        return await client.list_service_health_issues(params=params)

    @mcp.tool(
        name="get_service_health_issue",
        description="""get_service_health_issue: GET /admin/serviceAnnouncement/issues/{issue-id}



Get a specific service health issue.""",
        tags={"admin"},
    )
    async def get_service_health_issue(
        issue_id: str = Field(..., description="Parameter for issue-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_service_health_issue: GET /admin/serviceAnnouncement/issues/{issue-id}"""
        client = await get_client()
        return await client.get_service_health_issue(issue_id=issue_id, params=params)

    @mcp.tool(
        name="list_service_update_messages",
        description="""list_service_update_messages: GET /admin/serviceAnnouncement/messages



List service update messages (message center posts) for the tenant.""",
        tags={"admin"},
    )
    async def list_service_update_messages(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_service_update_messages: GET /admin/serviceAnnouncement/messages"""
        client = await get_client()
        return await client.list_service_update_messages(params=params)

    @mcp.tool(
        name="get_service_update_message",
        description="""get_service_update_message: GET /admin/serviceAnnouncement/messages/{message-id}



Get a specific service update message.""",
        tags={"admin"},
    )
    async def get_service_update_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_service_update_message: GET /admin/serviceAnnouncement/messages/{message-id}"""
        client = await get_client()
        return await client.get_service_update_message(
            message_id=message_id, params=params
        )

    @mcp.tool(
        name="get_admin_sharepoint",
        description="""get_admin_sharepoint: GET /admin/sharepoint



Get SharePoint admin settings for the tenant.""",
        tags={"admin", "sites"},
    )
    async def get_admin_sharepoint(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_admin_sharepoint: GET /admin/sharepoint"""
        client = await get_client()
        return await client.get_admin_sharepoint(params=params)

    @mcp.tool(
        name="update_admin_sharepoint",
        description="""update_admin_sharepoint: PATCH /admin/sharepoint



Update SharePoint admin settings for the tenant.""",
        tags={"admin", "sites"},
    )
    async def update_admin_sharepoint(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_admin_sharepoint: PATCH /admin/sharepoint"""
        client = await get_client()
        return await client.update_admin_sharepoint(data=data, params=params)

    @mcp.tool(
        name="list_delegated_admin_relationships",
        description="""list_delegated_admin_relationships: GET /tenantRelationships/delegatedAdminRelationships



List delegated admin relationships.""",
        tags={"admin"},
    )
    async def list_delegated_admin_relationships(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_delegated_admin_relationships: GET /tenantRelationships/delegatedAdminRelationships"""
        client = await get_client()
        return await client.list_delegated_admin_relationships(params=params)

    @mcp.tool(
        name="get_delegated_admin_relationship",
        description="""get_delegated_admin_relationship: GET /tenantRelationships/delegatedAdminRelationships/{id}



Get a specific delegated admin relationship.""",
        tags={"admin"},
    )
    async def get_delegated_admin_relationship(
        rel_id: str = Field(..., description="Delegated admin relationship ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_delegated_admin_relationship: GET /tenantRelationships/delegatedAdminRelationships/{id}"""
        client = await get_client()
        return await client.get_delegated_admin_relationship(
            rel_id=rel_id, params=params
        )


def register_organization_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_organization",
        description="""list_organization: GET /organization



Get the properties and relationships of the currently authenticated organization.""",
        tags={"organization"},
    )
    async def list_organization(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_organization: GET /organization"""
        client = await get_client()
        return await client.list_organization(params=params)

    @mcp.tool(
        name="get_organization",
        description="""get_organization: GET /organization/{org-id}



Get a specific organization by ID.""",
        tags={"organization"},
    )
    async def get_organization(
        org_id: str = Field(..., description="Parameter for organization-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_organization: GET /organization/{org-id}"""
        client = await get_client()
        return await client.get_organization(org_id=org_id, params=params)

    @mcp.tool(
        name="update_organization",
        description="""update_organization: PATCH /organization/{org-id}



Update organization properties.""",
        tags={"organization"},
    )
    async def update_organization(
        org_id: str = Field(..., description="Parameter for organization-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_organization: PATCH /organization/{org-id}"""
        client = await get_client()
        return await client.update_organization(org_id=org_id, data=data, params=params)

    @mcp.tool(
        name="get_org_branding",
        description="""get_org_branding: GET /organization/{org-id}/branding



Get organization branding properties (sign-in page customization).""",
        tags={"organization"},
    )
    async def get_org_branding(
        org_id: str = Field(..., description="Parameter for organization-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_org_branding: GET /organization/{org-id}/branding"""
        client = await get_client()
        return await client.get_org_branding(org_id=org_id, params=params)

    @mcp.tool(
        name="update_org_branding",
        description="""update_org_branding: PATCH /organization/{org-id}/branding



Update organization branding properties.""",
        tags={"organization"},
    )
    async def update_org_branding(
        org_id: str = Field(..., description="Parameter for organization-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_org_branding: PATCH /organization/{org-id}/branding"""
        client = await get_client()
        return await client.update_org_branding(org_id=org_id, data=data, params=params)


def register_domains_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_domains",
        description="""list_domains: GET /domains



List domains associated with the tenant.""",
        tags={"domains"},
    )
    async def list_domains(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_domains: GET /domains"""
        client = await get_client()
        return await client.list_domains(params=params)

    @mcp.tool(
        name="get_domain",
        description="""get_domain: GET /domains/{domain-id}



Get properties of a specific domain.""",
        tags={"domains"},
    )
    async def get_domain(
        domain_id: str = Field(..., description="Domain name (e.g. 'contoso.com')"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_domain: GET /domains/{domain-id}"""
        client = await get_client()
        return await client.get_domain(domain_id=domain_id, params=params)

    @mcp.tool(
        name="create_domain",
        description="""create_domain: POST /domains



Add a domain to the tenant. Provide the domain name as 'id' in the request body.""",
        tags={"domains"},
    )
    async def create_domain(
        data: dict[str, Any] | None = Field(
            None, description="Request body data with 'id' (domain name)"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_domain: POST /domains"""
        client = await get_client()
        return await client.create_domain(data=data, params=params)

    @mcp.tool(
        name="delete_domain",
        description="""delete_domain: DELETE /domains/{domain-id}



Delete a domain from the tenant.""",
        tags={"domains"},
    )
    async def delete_domain(
        domain_id: str = Field(..., description="Domain name to delete"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_domain: DELETE /domains/{domain-id}"""
        if not await ctx_confirm_destructive(ctx, "delete domain"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_domain(domain_id=domain_id, params=params)

    @mcp.tool(
        name="verify_domain",
        description="""verify_domain: POST /domains/{domain-id}/verify



Verify ownership of a domain.""",
        tags={"domains"},
    )
    async def verify_domain(
        domain_id: str = Field(..., description="Domain name to verify"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """verify_domain: POST /domains/{domain-id}/verify"""
        client = await get_client()
        return await client.verify_domain(domain_id=domain_id, params=params)

    @mcp.tool(
        name="list_domain_service_configuration_records",
        description="""list_domain_service_configuration_records: GET /domains/{domain-id}/serviceConfigurationRecords



List DNS records required by the domain for Microsoft services.""",
        tags={"domains"},
    )
    async def list_domain_service_configuration_records(
        domain_id: str = Field(..., description="Domain name"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_domain_service_configuration_records: GET /domains/{domain-id}/serviceConfigurationRecords"""
        client = await get_client()
        return await client.list_domain_service_configuration_records(
            domain_id=domain_id, params=params
        )


def register_subscriptions_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_subscriptions",
        description="""list_subscriptions: GET /subscriptions



List active webhook subscriptions for change notifications.""",
        tags={"subscriptions"},
    )
    async def list_subscriptions(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_subscriptions: GET /subscriptions"""
        client = await get_client()
        return await client.list_subscriptions(params=params)

    @mcp.tool(
        name="get_subscription",
        description="""get_subscription: GET /subscriptions/{subscription-id}



Get a specific subscription.""",
        tags={"subscriptions"},
    )
    async def get_subscription(
        subscription_id: str = Field(..., description="Parameter for subscription-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_subscription: GET /subscriptions/{subscription-id}"""
        client = await get_client()
        return await client.get_subscription(
            subscription_id=subscription_id, params=params
        )

    @mcp.tool(
        name="create_subscription",
        description="""create_subscription: POST /subscriptions



Create a webhook subscription for change notifications. Required fields: changeType, notificationUrl, resource, expirationDateTime.""",
        tags={"subscriptions"},
    )
    async def create_subscription(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_subscription: POST /subscriptions"""
        client = await get_client()
        return await client.create_subscription(data=data, params=params)

    @mcp.tool(
        name="update_subscription",
        description="""update_subscription: PATCH /subscriptions/{subscription-id}



Renew a subscription by extending its expiration time.""",
        tags={"subscriptions"},
    )
    async def update_subscription(
        subscription_id: str = Field(..., description="Parameter for subscription-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_subscription: PATCH /subscriptions/{subscription-id}"""
        client = await get_client()
        return await client.update_subscription(
            subscription_id=subscription_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_subscription",
        description="""delete_subscription: DELETE /subscriptions/{subscription-id}



Delete a webhook subscription.""",
        tags={"subscriptions"},
    )
    async def delete_subscription(
        subscription_id: str = Field(..., description="Parameter for subscription-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_subscription: DELETE /subscriptions/{subscription-id}"""
        if not await ctx_confirm_destructive(ctx, "delete subscription"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_subscription(
            subscription_id=subscription_id, params=params
        )


def register_communications_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_online_meetings",
        description="""list_online_meetings: GET /me/onlineMeetings



List online meetings for the current user. Returns meeting details including subject, join URL, start/end time, and participants.""",
        tags={"communications"},
    )
    async def list_online_meetings(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_online_meetings: GET /me/onlineMeetings"""
        client = await get_client()
        return await client.list_online_meetings(params=params)

    @mcp.tool(
        name="get_online_meeting",
        description="""get_online_meeting: GET /me/onlineMeetings/{onlineMeeting-id}



Get a specific online meeting by ID. Returns full meeting details including join information, audio conferencing, and lobby settings.""",
        tags={"communications"},
    )
    async def get_online_meeting(
        meeting_id: str = Field(..., description="Parameter for onlineMeeting-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_online_meeting: GET /me/onlineMeetings/{onlineMeeting-id}"""
        client = await get_client()
        return await client.get_online_meeting(meeting_id=meeting_id, params=params)

    @mcp.tool(
        name="create_online_meeting",
        description="""create_online_meeting: POST /me/onlineMeetings



Create a new online meeting. Provide subject, startDateTime, endDateTime, and optional lobby bypass settings.""",
        tags={"communications"},
    )
    async def create_online_meeting(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_online_meeting: POST /me/onlineMeetings"""
        client = await get_client()
        return await client.create_online_meeting(data=data, params=params)

    @mcp.tool(
        name="update_online_meeting",
        description="""update_online_meeting: PATCH /me/onlineMeetings/{onlineMeeting-id}



Update an existing online meeting. Modify subject, times, or lobby settings.""",
        tags={"communications"},
    )
    async def update_online_meeting(
        meeting_id: str = Field(..., description="Parameter for onlineMeeting-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_online_meeting: PATCH /me/onlineMeetings/{onlineMeeting-id}"""
        client = await get_client()
        return await client.update_online_meeting(
            meeting_id=meeting_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_online_meeting",
        description="""delete_online_meeting: DELETE /me/onlineMeetings/{onlineMeeting-id}



Delete an online meeting.""",
        tags={"communications"},
    )
    async def delete_online_meeting(
        meeting_id: str = Field(..., description="Parameter for onlineMeeting-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_online_meeting: DELETE /me/onlineMeetings/{onlineMeeting-id}"""
        if not await ctx_confirm_destructive(ctx, "delete online meeting"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_online_meeting(meeting_id=meeting_id, params=params)

    @mcp.tool(
        name="list_call_records",
        description="""list_call_records: GET /communications/callRecords



List call records. Returns information about calls and meetings including participants, modalities, and duration.""",
        tags={"communications"},
    )
    async def list_call_records(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_call_records: GET /communications/callRecords"""
        client = await get_client()
        return await client.list_call_records(params=params)

    @mcp.tool(
        name="get_call_record",
        description="""get_call_record: GET /communications/callRecords/{callRecord-id}



Get a specific call record by ID. Returns detailed call information including sessions and segments.""",
        tags={"communications"},
    )
    async def get_call_record(
        call_id: str = Field(..., description="Parameter for callRecord-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_call_record: GET /communications/callRecords/{callRecord-id}"""
        client = await get_client()
        return await client.get_call_record(call_id=call_id, params=params)

    @mcp.tool(
        name="list_presences",
        description="""list_presences: GET /communications/presences



List presence information for multiple users. Returns availability and activity status.""",
        tags={"communications"},
    )
    async def list_presences(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_presences: GET /communications/presences"""
        client = await get_client()
        return await client.list_presences(params=params)

    @mcp.tool(
        name="get_presence",
        description="""get_presence: GET /communications/presences/{presence-id}



Get presence for a specific user by user ID. Returns availability (Available, Busy, Away, etc.) and activity.""",
        tags={"communications"},
    )
    async def get_presence(
        user_id: str = Field(..., description="User ID to get presence for"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_presence: GET /communications/presences/{presence-id}"""
        client = await get_client()
        return await client.get_presence(user_id=user_id, params=params)

    @mcp.tool(
        name="get_my_presence",
        description="""get_my_presence: GET /me/presence



Get current user's presence status including availability and activity.""",
        tags={"communications"},
    )
    async def get_my_presence(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_my_presence: GET /me/presence"""
        client = await get_client()
        return await client.get_my_presence(params=params)


def register_identity_tools(mcp: FastMCP):
    @mcp.tool(
        name="create_invitation",
        description="""create_invitation: POST /invitations



Create an invitation for an external / guest user. Provide invitedUserEmailAddress and inviteRedirectUrl. Optionally set invitedUserDisplayName and sendInvitationMessage.""",
        tags={"identity"},
    )
    async def create_invitation(
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_invitation: POST /invitations"""
        client = await get_client()
        return await client.create_invitation(data=data, params=params)

    @mcp.tool(
        name="list_conditional_access_policies",
        description="""list_conditional_access_policies: GET /identity/conditionalAccess/policies



List conditional access policies.""",
        tags={"identity"},
    )
    async def list_conditional_access_policies(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_conditional_access_policies: GET /identity/conditionalAccess/policies"""
        client = await get_client()
        return await client.list_conditional_access_policies(params=params)

    @mcp.tool(
        name="get_conditional_access_policy",
        description="""get_conditional_access_policy: GET /identity/conditionalAccess/policies/{id}



Get a specific conditional access policy.""",
        tags={"identity"},
    )
    async def get_conditional_access_policy(
        policy_id: str = Field(..., description="Conditional access policy ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_conditional_access_policy: GET /identity/conditionalAccess/policies/{id}"""
        client = await get_client()
        return await client.get_conditional_access_policy(
            policy_id=policy_id, params=params
        )

    @mcp.tool(
        name="create_conditional_access_policy",
        description="""create_conditional_access_policy: POST /identity/conditionalAccess/policies



Create a conditional access policy.""",
        tags={"identity"},
    )
    async def create_conditional_access_policy(
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName, state, conditions, etc."
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_conditional_access_policy: POST /identity/conditionalAccess/policies"""
        client = await get_client()
        return await client.create_conditional_access_policy(data=data, params=params)

    @mcp.tool(
        name="update_conditional_access_policy",
        description="""update_conditional_access_policy: PATCH /identity/conditionalAccess/policies/{id}



Update a conditional access policy.""",
        tags={"identity"},
    )
    async def update_conditional_access_policy(
        policy_id: str = Field(..., description="Conditional access policy ID"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_conditional_access_policy: PATCH /identity/conditionalAccess/policies/{id}"""
        client = await get_client()
        return await client.update_conditional_access_policy(
            policy_id=policy_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_conditional_access_policy",
        description="""delete_conditional_access_policy: DELETE /identity/conditionalAccess/policies/{id}



Delete a conditional access policy.""",
        tags={"identity"},
    )
    async def delete_conditional_access_policy(
        policy_id: str = Field(..., description="Conditional access policy ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_conditional_access_policy: DELETE /identity/conditionalAccess/policies/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete conditional access policy"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_conditional_access_policy(
            policy_id=policy_id, params=params
        )

    @mcp.tool(
        name="list_access_reviews",
        description="""list_access_reviews: GET /identityGovernance/accessReviewDefinitions



List access review schedule definitions.""",
        tags={"identity"},
    )
    async def list_access_reviews(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_access_reviews: GET /identityGovernance/accessReviewDefinitions"""
        client = await get_client()
        return await client.list_access_reviews(params=params)

    @mcp.tool(
        name="get_access_review",
        description="""get_access_review: GET /identityGovernance/accessReviewDefinitions/{id}



Get a specific access review definition.""",
        tags={"identity"},
    )
    async def get_access_review(
        review_id: str = Field(..., description="Access review schedule definition ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_access_review: GET /identityGovernance/accessReviewDefinitions/{id}"""
        client = await get_client()
        return await client.get_access_review(review_id=review_id, params=params)

    @mcp.tool(
        name="list_entitlement_access_packages",
        description="""list_entitlement_access_packages: GET /identityGovernance/entitlementManagement/accessPackages



List entitlement management access packages.""",
        tags={"identity"},
    )
    async def list_entitlement_access_packages(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_entitlement_access_packages: GET /identityGovernance/entitlementManagement/accessPackages"""
        client = await get_client()
        return await client.list_entitlement_access_packages(params=params)

    @mcp.tool(
        name="list_lifecycle_workflows",
        description="""list_lifecycle_workflows: GET /identityGovernance/lifecycleWorkflows/workflows



List lifecycle management workflows.""",
        tags={"identity"},
    )
    async def list_lifecycle_workflows(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_lifecycle_workflows: GET /identityGovernance/lifecycleWorkflows/workflows"""
        client = await get_client()
        return await client.list_lifecycle_workflows(params=params)


def register_security_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_security_alerts",
        description="""list_security_alerts: GET /security/alerts_v2



List security alerts. Returns alert details including severity, status, and detected threats.""",
        tags={"security"},
    )
    async def list_security_alerts(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_security_alerts: GET /security/alerts_v2"""
        client = await get_client()
        return await client.list_security_alerts(params=params)

    @mcp.tool(
        name="get_security_alert",
        description="""get_security_alert: GET /security/alerts_v2/{alert-id}



Get a specific security alert by ID.""",
        tags={"security"},
    )
    async def get_security_alert(
        alert_id: str = Field(..., description="Parameter for alert-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_security_alert: GET /security/alerts_v2/{alert-id}"""
        client = await get_client()
        return await client.get_security_alert(alert_id=alert_id, params=params)

    @mcp.tool(
        name="update_security_alert",
        description="""update_security_alert: PATCH /security/alerts_v2/{alert-id}



Update a security alert. Change status, assign to user, set classification/determination.""",
        tags={"security"},
    )
    async def update_security_alert(
        alert_id: str = Field(..., description="Parameter for alert-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_security_alert: PATCH /security/alerts_v2/{alert-id}"""
        client = await get_client()
        return await client.update_security_alert(
            alert_id=alert_id, data=data, params=params
        )

    @mcp.tool(
        name="list_security_incidents",
        description="""list_security_incidents: GET /security/incidents



List security incidents. Returns correlated alerts grouped into incidents.""",
        tags={"security"},
    )
    async def list_security_incidents(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_security_incidents: GET /security/incidents"""
        client = await get_client()
        return await client.list_security_incidents(params=params)

    @mcp.tool(
        name="get_security_incident",
        description="""get_security_incident: GET /security/incidents/{incident-id}



Get a specific security incident by ID.""",
        tags={"security"},
    )
    async def get_security_incident(
        incident_id: str = Field(..., description="Parameter for incident-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_security_incident: GET /security/incidents/{incident-id}"""
        client = await get_client()
        return await client.get_security_incident(
            incident_id=incident_id, params=params
        )

    @mcp.tool(
        name="update_security_incident",
        description="""update_security_incident: PATCH /security/incidents/{incident-id}



Update a security incident. Change status, assign, classify.""",
        tags={"security"},
    )
    async def update_security_incident(
        incident_id: str = Field(..., description="Parameter for incident-id"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_security_incident: PATCH /security/incidents/{incident-id}"""
        client = await get_client()
        return await client.update_security_incident(
            incident_id=incident_id, data=data, params=params
        )

    @mcp.tool(
        name="list_secure_scores",
        description="""list_secure_scores: GET /security/secureScores



List tenant secure scores over time.""",
        tags={"security"},
    )
    async def list_secure_scores(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_secure_scores: GET /security/secureScores"""
        client = await get_client()
        return await client.list_secure_scores(params=params)

    @mcp.tool(
        name="list_threat_intelligence_hosts",
        description="""list_threat_intelligence_hosts: GET /security/threatIntelligence/hosts



List threat intelligence hosts.""",
        tags={"security"},
    )
    async def list_threat_intelligence_hosts(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_threat_intelligence_hosts: GET /security/threatIntelligence/hosts"""
        client = await get_client()
        return await client.list_threat_intelligence_hosts(params=params)

    @mcp.tool(
        name="get_threat_intelligence_host",
        description="""get_threat_intelligence_host: GET /security/threatIntelligence/hosts/{host-id}



Get a specific threat intelligence host.""",
        tags={"security"},
    )
    async def get_threat_intelligence_host(
        host_id: str = Field(..., description="Parameter for host-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_threat_intelligence_host: GET /security/threatIntelligence/hosts/{host-id}"""
        client = await get_client()
        return await client.get_threat_intelligence_host(host_id=host_id, params=params)

    @mcp.tool(
        name="run_hunting_query",
        description="""run_hunting_query: POST /security/runHuntingQuery



Run an advanced hunting query using Kusto Query Language (KQL).""",
        tags={"security"},
    )
    async def run_hunting_query(
        data: dict[str, Any] | None = Field(
            None, description="Request body data with 'query' field"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        await ctx_progress(ctx, 0, 100)
        """run_hunting_query: POST /security/runHuntingQuery"""
        client = await get_client()
        await ctx_progress(ctx, 100, 100)
        return await client.run_hunting_query(data=data, params=params)

    @mcp.tool(
        name="list_risk_detections",
        description="""list_risk_detections: GET /identityProtection/riskDetections



List risk detections (sign-in anomalies, leaked credentials, etc.).""",
        tags={"security"},
    )
    async def list_risk_detections(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_risk_detections: GET /identityProtection/riskDetections"""
        client = await get_client()
        return await client.list_risk_detections(params=params)

    @mcp.tool(
        name="get_risk_detection",
        description="""get_risk_detection: GET /identityProtection/riskDetections/{id}



Get a specific risk detection.""",
        tags={"security"},
    )
    async def get_risk_detection(
        risk_id: str = Field(..., description="Risk detection ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_risk_detection: GET /identityProtection/riskDetections/{id}"""
        client = await get_client()
        return await client.get_risk_detection(risk_id=risk_id, params=params)

    @mcp.tool(
        name="list_risky_users",
        description="""list_risky_users: GET /identityProtection/riskyUsers



List users flagged as risky.""",
        tags={"security"},
    )
    async def list_risky_users(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_risky_users: GET /identityProtection/riskyUsers"""
        client = await get_client()
        return await client.list_risky_users(params=params)

    @mcp.tool(
        name="get_risky_user",
        description="""get_risky_user: GET /identityProtection/riskyUsers/{id}



Get a specific risky user.""",
        tags={"security"},
    )
    async def get_risky_user(
        user_id: str = Field(..., description="Risky user ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_risky_user: GET /identityProtection/riskyUsers/{id}"""
        client = await get_client()
        return await client.get_risky_user(user_id=user_id, params=params)

    @mcp.tool(
        name="dismiss_risky_user",
        description="""dismiss_risky_user: POST /identityProtection/riskyUsers/dismiss



Dismiss a risky user (mark as safe).""",
        tags={"security"},
    )
    async def dismiss_risky_user(
        user_id: str = Field(..., description="User ID to dismiss"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """dismiss_risky_user: POST /identityProtection/riskyUsers/dismiss"""
        if not await ctx_confirm_destructive(ctx, "dismiss risky user"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.dismiss_risky_user(user_id=user_id, params=params)

    @mcp.tool(
        name="list_sensitivity_labels",
        description="""list_sensitivity_labels: GET /informationProtection/policy/labels



List sensitivity labels.""",
        tags={"security"},
    )
    async def list_sensitivity_labels(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_sensitivity_labels: GET /informationProtection/policy/labels"""
        client = await get_client()
        return await client.list_sensitivity_labels(params=params)

    @mcp.tool(
        name="get_sensitivity_label",
        description="""get_sensitivity_label: GET /informationProtection/policy/labels/{id}



Get a specific sensitivity label.""",
        tags={"security"},
    )
    async def get_sensitivity_label(
        label_id: str = Field(..., description="Sensitivity label ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sensitivity_label: GET /informationProtection/policy/labels/{id}"""
        client = await get_client()
        return await client.get_sensitivity_label(label_id=label_id, params=params)


def register_audit_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_directory_audits",
        description="""list_directory_audits: GET /auditLogs/directoryAudits



List directory audit log entries. Shows changes to directory objects (users, groups, apps).""",
        tags={"audit"},
    )
    async def list_directory_audits(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_directory_audits: GET /auditLogs/directoryAudits"""
        client = await get_client()
        return await client.list_directory_audits(params=params)

    @mcp.tool(
        name="get_directory_audit",
        description="""get_directory_audit: GET /auditLogs/directoryAudits/{directoryAudit-id}



Get a specific directory audit entry.""",
        tags={"audit"},
    )
    async def get_directory_audit(
        audit_id: str = Field(..., description="Parameter for directoryAudit-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_directory_audit: GET /auditLogs/directoryAudits/{directoryAudit-id}"""
        client = await get_client()
        return await client.get_directory_audit(audit_id=audit_id, params=params)

    @mcp.tool(
        name="list_sign_in_logs",
        description="""list_sign_in_logs: GET /auditLogs/signIns



List sign-in activity logs. Shows user sign-in events with details.""",
        tags={"audit"},
    )
    async def list_sign_in_logs(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_sign_in_logs: GET /auditLogs/signIns"""
        client = await get_client()
        return await client.list_sign_in_logs(params=params)

    @mcp.tool(
        name="get_sign_in_log",
        description="""get_sign_in_log: GET /auditLogs/signIns/{signIn-id}



Get a specific sign-in log entry.""",
        tags={"audit"},
    )
    async def get_sign_in_log(
        sign_in_id: str = Field(..., description="Parameter for signIn-id"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sign_in_log: GET /auditLogs/signIns/{signIn-id}"""
        client = await get_client()
        return await client.get_sign_in_log(sign_in_id=sign_in_id, params=params)

    @mcp.tool(
        name="list_provisioning_logs",
        description="""list_provisioning_logs: GET /auditLogs/provisioning



List provisioning logs. Shows automated user/group provisioning events.""",
        tags={"audit"},
    )
    async def list_provisioning_logs(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_provisioning_logs: GET /auditLogs/provisioning"""
        client = await get_client()
        return await client.list_provisioning_logs(params=params)


def register_reports_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_email_activity_report",
        description="""get_email_activity_report: GET /reports/getEmailActivityUserDetail



Get email activity user detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_email_activity_report(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_email_activity_report: GET /reports/getEmailActivityUserDetail"""
        client = await get_client()
        return await client.get_email_activity_report(period=period, params=params)

    @mcp.tool(
        name="get_mailbox_usage_report",
        description="""get_mailbox_usage_report: GET /reports/getMailboxUsageDetail



Get mailbox usage detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_mailbox_usage_report(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_mailbox_usage_report: GET /reports/getMailboxUsageDetail"""
        client = await get_client()
        return await client.get_mailbox_usage_report(period=period, params=params)

    @mcp.tool(
        name="get_office365_active_users",
        description="""get_office365_active_users: GET /reports/getOffice365ActiveUserDetail



Get Office 365 active user detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_office365_active_users(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_office365_active_users: GET /reports/getOffice365ActiveUserDetail"""
        client = await get_client()
        return await client.get_office365_active_users(period=period, params=params)

    @mcp.tool(
        name="get_sharepoint_activity_report",
        description="""get_sharepoint_activity_report: GET /reports/getSharePointActivityUserDetail



Get SharePoint activity user detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_sharepoint_activity_report(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_sharepoint_activity_report: GET /reports/getSharePointActivityUserDetail"""
        client = await get_client()
        return await client.get_sharepoint_activity_report(period=period, params=params)

    @mcp.tool(
        name="get_teams_user_activity",
        description="""get_teams_user_activity: GET /reports/getTeamsUserActivityUserDetail



Get Teams user activity detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_teams_user_activity(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_teams_user_activity: GET /reports/getTeamsUserActivityUserDetail"""
        client = await get_client()
        return await client.get_teams_user_activity(period=period, params=params)

    @mcp.tool(
        name="get_onedrive_usage_report",
        description="""get_onedrive_usage_report: GET /reports/getOneDriveUsageAccountDetail



Get OneDrive usage account detail report. Period: D7, D30, D90, D180.""",
        tags={"reports"},
    )
    async def get_onedrive_usage_report(
        period: str = Field("D7", description="Report period: D7, D30, D90, or D180"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_onedrive_usage_report: GET /reports/getOneDriveUsageAccountDetail"""
        client = await get_client()
        return await client.get_onedrive_usage_report(period=period, params=params)


def register_applications_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_applications",
        description="""list_applications: GET /applications



List app registrations in the tenant.""",
        tags={"applications"},
    )
    async def list_applications(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_applications: GET /applications"""
        client = await get_client()
        return await client.list_applications(params=params)

    @mcp.tool(
        name="get_application",
        description="""get_application: GET /applications/{id}



Get a specific app registration.""",
        tags={"applications"},
    )
    async def get_application(
        app_id: str = Field(..., description="Application object ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_application: GET /applications/{id}"""
        client = await get_client()
        return await client.get_application(app_id=app_id, params=params)

    @mcp.tool(
        name="create_application",
        description="""create_application: POST /applications



Create an app registration.""",
        tags={"applications"},
    )
    async def create_application(
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName, signInAudience, etc."
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_application: POST /applications"""
        client = await get_client()
        return await client.create_application(data=data, params=params)

    @mcp.tool(
        name="update_application",
        description="""update_application: PATCH /applications/{id}



Update an app registration.""",
        tags={"applications"},
    )
    async def update_application(
        app_id: str = Field(..., description="Application object ID"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_application: PATCH /applications/{id}"""
        client = await get_client()
        return await client.update_application(app_id=app_id, data=data, params=params)

    @mcp.tool(
        name="delete_application",
        description="""delete_application: DELETE /applications/{id}



Delete an app registration.""",
        tags={"applications"},
    )
    async def delete_application(
        app_id: str = Field(..., description="Application object ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_application: DELETE /applications/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete application"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_application(app_id=app_id, params=params)

    @mcp.tool(
        name="add_application_password",
        description="""add_application_password: POST /applications/{id}/addPassword



Add a password credential (client secret) to an app.""",
        tags={"applications"},
    )
    async def add_application_password(
        app_id: str = Field(..., description="Application object ID"),
        data: dict[str, Any] | None = Field(
            None, description="Request body with optional displayName"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """add_application_password: POST /applications/{id}/addPassword"""
        client = await get_client()
        return await client.add_application_password(
            app_id=app_id, data=data, params=params
        )

    @mcp.tool(
        name="remove_application_password",
        description="""remove_application_password: POST /applications/{id}/removePassword



Remove a password credential from an app.""",
        tags={"applications"},
    )
    async def remove_application_password(
        app_id: str = Field(..., description="Application object ID"),
        data: dict[str, Any] | None = Field(
            None, description="Request body with keyId (UUID of the credential)"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """remove_application_password: POST /applications/{id}/removePassword"""
        if not await ctx_confirm_destructive(ctx, "remove application password"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.remove_application_password(
            app_id=app_id, data=data, params=params
        )

    @mcp.tool(
        name="list_service_principals",
        description="""list_service_principals: GET /servicePrincipals



List service principals (enterprise apps).""",
        tags={"applications"},
    )
    async def list_service_principals(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_service_principals: GET /servicePrincipals"""
        client = await get_client()
        return await client.list_service_principals(params=params)

    @mcp.tool(
        name="get_service_principal",
        description="""get_service_principal: GET /servicePrincipals/{id}



Get a specific service principal.""",
        tags={"applications"},
    )
    async def get_service_principal(
        sp_id: str = Field(..., description="Service principal ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_service_principal: GET /servicePrincipals/{id}"""
        client = await get_client()
        return await client.get_service_principal(sp_id=sp_id, params=params)

    @mcp.tool(
        name="create_service_principal",
        description="""create_service_principal: POST /servicePrincipals



Create a service principal for an app.""",
        tags={"applications"},
    )
    async def create_service_principal(
        data: dict[str, Any] | None = Field(
            None, description="Request body with appId"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_service_principal: POST /servicePrincipals"""
        client = await get_client()
        return await client.create_service_principal(data=data, params=params)

    @mcp.tool(
        name="update_service_principal",
        description="""update_service_principal: PATCH /servicePrincipals/{id}



Update a service principal.""",
        tags={"applications"},
    )
    async def update_service_principal(
        sp_id: str = Field(..., description="Service principal ID"),
        data: dict[str, Any] | None = Field(None, description="Request body data"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_service_principal: PATCH /servicePrincipals/{id}"""
        client = await get_client()
        return await client.update_service_principal(
            sp_id=sp_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_service_principal",
        description="""delete_service_principal: DELETE /servicePrincipals/{id}



Delete a service principal.""",
        tags={"applications"},
    )
    async def delete_service_principal(
        sp_id: str = Field(..., description="Service principal ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_service_principal: DELETE /servicePrincipals/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete service principal"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_service_principal(sp_id=sp_id, params=params)


def register_directory_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_directory_objects",
        description="""list_directory_objects: GET /directoryObjects



List directory objects.""",
        tags={"directory"},
    )
    async def list_directory_objects(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_directory_objects: GET /directoryObjects"""
        client = await get_client()
        return await client.list_directory_objects(params=params)

    @mcp.tool(
        name="get_directory_object",
        description="""get_directory_object: GET /directoryObjects/{id}



Get a specific directory object.""",
        tags={"directory"},
    )
    async def get_directory_object(
        object_id: str = Field(..., description="Directory object ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_directory_object: GET /directoryObjects/{id}"""
        client = await get_client()
        return await client.get_directory_object(object_id=object_id, params=params)

    @mcp.tool(
        name="list_directory_roles",
        description="""list_directory_roles: GET /directoryRoles



List activated directory roles.""",
        tags={"directory"},
    )
    async def list_directory_roles(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_directory_roles: GET /directoryRoles"""
        client = await get_client()
        return await client.list_directory_roles(params=params)

    @mcp.tool(
        name="get_directory_role",
        description="""get_directory_role: GET /directoryRoles/{id}



Get a specific activated directory role.""",
        tags={"directory"},
    )
    async def get_directory_role(
        role_id: str = Field(..., description="Directory role ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_directory_role: GET /directoryRoles/{id}"""
        client = await get_client()
        return await client.get_directory_role(role_id=role_id, params=params)

    @mcp.tool(
        name="list_directory_role_templates",
        description="""list_directory_role_templates: GET /directoryRoleTemplates



List all directory role templates (built-in role definitions).""",
        tags={"directory"},
    )
    async def list_directory_role_templates(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_directory_role_templates: GET /directoryRoleTemplates"""
        client = await get_client()
        return await client.list_directory_role_templates(params=params)

    @mcp.tool(
        name="list_deleted_items",
        description="""list_deleted_items: GET /directory/deletedItems



List recently deleted directory items (users, groups, apps).""",
        tags={"directory"},
    )
    async def list_deleted_items(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_deleted_items: GET /directory/deletedItems"""
        if not await ctx_confirm_destructive(ctx, "list deleted items"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.list_deleted_items(params=params)

    @mcp.tool(
        name="restore_deleted_item",
        description="""restore_deleted_item: POST /directory/deletedItems/{id}/restore



Restore a recently deleted directory item.""",
        tags={"directory"},
    )
    async def restore_deleted_item(
        object_id: str = Field(..., description="Deleted object ID to restore"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """restore_deleted_item: POST /directory/deletedItems/{id}/restore"""
        if not await ctx_confirm_destructive(ctx, "restore deleted item"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.restore_deleted_item(object_id=object_id, params=params)

    @mcp.tool(
        name="list_role_definitions",
        description="""list_role_definitions: GET /roleManagement/directory/roleDefinitions



List RBAC directory role definitions.""",
        tags={"directory"},
    )
    async def list_role_definitions(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_role_definitions: GET /roleManagement/directory/roleDefinitions"""
        client = await get_client()
        return await client.list_role_definitions(params=params)

    @mcp.tool(
        name="get_role_definition",
        description="""get_role_definition: GET /roleManagement/directory/roleDefinitions/{id}



Get a specific RBAC role definition.""",
        tags={"directory"},
    )
    async def get_role_definition(
        role_id: str = Field(..., description="Role definition ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_role_definition: GET /roleManagement/directory/roleDefinitions/{id}"""
        client = await get_client()
        return await client.get_role_definition(role_id=role_id, params=params)

    @mcp.tool(
        name="list_role_assignments",
        description="""list_role_assignments: GET /roleManagement/directory/roleAssignments



List RBAC directory role assignments.""",
        tags={"directory"},
    )
    async def list_role_assignments(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_role_assignments: GET /roleManagement/directory/roleAssignments"""
        client = await get_client()
        return await client.list_role_assignments(params=params)

    @mcp.tool(
        name="get_role_assignment",
        description="""get_role_assignment: GET /roleManagement/directory/roleAssignments/{id}



Get a specific RBAC role assignment.""",
        tags={"directory"},
    )
    async def get_role_assignment(
        assignment_id: str = Field(..., description="Role assignment ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_role_assignment: GET /roleManagement/directory/roleAssignments/{id}"""
        client = await get_client()
        return await client.get_role_assignment(
            assignment_id=assignment_id, params=params
        )

    @mcp.tool(
        name="create_role_assignment",
        description="""create_role_assignment: POST /roleManagement/directory/roleAssignments



Create a new RBAC role assignment.""",
        tags={"directory"},
    )
    async def create_role_assignment(
        data: dict[str, Any] | None = Field(
            None,
            description="Request body with roleDefinitionId, principalId, directoryScopeId",
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_role_assignment: POST /roleManagement/directory/roleAssignments"""
        client = await get_client()
        return await client.create_role_assignment(data=data, params=params)


def register_policies_tools(mcp: FastMCP):
    @mcp.tool(
        name="get_authorization_policy",
        description="""get_authorization_policy: GET /policies/authorizationPolicy



Get the tenant authorization policy.""",
        tags={"policies"},
    )
    async def get_authorization_policy(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_authorization_policy: GET /policies/authorizationPolicy"""
        client = await get_client()
        return await client.get_authorization_policy(params=params)

    @mcp.tool(
        name="list_token_lifetime_policies",
        description="""list_token_lifetime_policies: GET /policies/tokenLifetimePolicies



List token lifetime policies.""",
        tags={"policies"},
    )
    async def list_token_lifetime_policies(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_token_lifetime_policies: GET /policies/tokenLifetimePolicies"""
        client = await get_client()
        return await client.list_token_lifetime_policies(params=params)

    @mcp.tool(
        name="list_token_issuance_policies",
        description="""list_token_issuance_policies: GET /policies/tokenIssuancePolicies



List token issuance policies.""",
        tags={"policies"},
    )
    async def list_token_issuance_policies(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_token_issuance_policies: GET /policies/tokenIssuancePolicies"""
        client = await get_client()
        return await client.list_token_issuance_policies(params=params)

    @mcp.tool(
        name="list_permission_grant_policies",
        description="""list_permission_grant_policies: GET /policies/permissionGrantPolicies



List permission grant policies.""",
        tags={"policies"},
    )
    async def list_permission_grant_policies(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_permission_grant_policies: GET /policies/permissionGrantPolicies"""
        client = await get_client()
        return await client.list_permission_grant_policies(params=params)

    @mcp.tool(
        name="get_admin_consent_policy",
        description="""get_admin_consent_policy: GET /policies/adminConsentRequestPolicy



Get the admin consent request policy.""",
        tags={"policies"},
    )
    async def get_admin_consent_policy(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_admin_consent_policy: GET /policies/adminConsentRequestPolicy"""
        client = await get_client()
        return await client.get_admin_consent_policy(params=params)


def register_devices_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_devices",
        description="""list_devices: GET /devices



List devices registered in the directory.""",
        tags={"devices"},
    )
    async def list_devices(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_devices: GET /devices"""
        client = await get_client()
        return await client.list_devices(params=params)

    @mcp.tool(
        name="get_device",
        description="""get_device: GET /devices/{id}



Get a specific device.""",
        tags={"devices"},
    )
    async def get_device(
        device_id: str = Field(..., description="Device ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_device: GET /devices/{id}"""
        client = await get_client()
        return await client.get_device(device_id=device_id, params=params)

    @mcp.tool(
        name="delete_device",
        description="""delete_device: DELETE /devices/{id}



Delete a device.""",
        tags={"devices"},
    )
    async def delete_device(
        device_id: str = Field(..., description="Device ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_device: DELETE /devices/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete device"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_device(device_id=device_id, params=params)

    @mcp.tool(
        name="list_managed_devices",
        description="""list_managed_devices: GET /deviceManagement/managedDevices



List Intune managed devices.""",
        tags={"devices"},
    )
    async def list_managed_devices(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_managed_devices: GET /deviceManagement/managedDevices"""
        client = await get_client()
        return await client.list_managed_devices(params=params)

    @mcp.tool(
        name="get_managed_device",
        description="""get_managed_device: GET /deviceManagement/managedDevices/{id}



Get a specific managed device.""",
        tags={"devices"},
    )
    async def get_managed_device(
        device_id: str = Field(..., description="Managed device ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_managed_device: GET /deviceManagement/managedDevices/{id}"""
        client = await get_client()
        return await client.get_managed_device(device_id=device_id, params=params)

    @mcp.tool(
        name="list_device_compliance_policies",
        description="""list_device_compliance_policies: GET /deviceManagement/deviceCompliancePolicies



List device compliance policies.""",
        tags={"devices"},
    )
    async def list_device_compliance_policies(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_device_compliance_policies: GET /deviceManagement/deviceCompliancePolicies"""
        client = await get_client()
        return await client.list_device_compliance_policies(params=params)

    @mcp.tool(
        name="list_device_configurations",
        description="""list_device_configurations: GET /deviceManagement/deviceConfigurations



List device configuration profiles.""",
        tags={"devices"},
    )
    async def list_device_configurations(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_device_configurations: GET /deviceManagement/deviceConfigurations"""
        client = await get_client()
        return await client.list_device_configurations(params=params)

    @mcp.tool(
        name="wipe_managed_device",
        description="""wipe_managed_device: POST /deviceManagement/managedDevices/{id}/wipe



Wipe a managed device (factory reset).""",
        tags={"devices"},
    )
    async def wipe_managed_device(
        device_id: str = Field(..., description="Managed device ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """wipe_managed_device: POST /deviceManagement/managedDevices/{id}/wipe"""
        if not await ctx_confirm_destructive(ctx, "wipe managed device"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.wipe_managed_device(device_id=device_id, params=params)

    @mcp.tool(
        name="retire_managed_device",
        description="""retire_managed_device: POST /deviceManagement/managedDevices/{id}/retire



Retire a managed device (remove company data).""",
        tags={"devices"},
    )
    async def retire_managed_device(
        device_id: str = Field(..., description="Managed device ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """retire_managed_device: POST /deviceManagement/managedDevices/{id}/retire"""
        client = await get_client()
        return await client.retire_managed_device(device_id=device_id, params=params)


def register_education_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_education_classes",
        description="""list_education_classes: GET /education/classes



List education classes.""",
        tags={"education"},
    )
    async def list_education_classes(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_education_classes: GET /education/classes"""
        client = await get_client()
        return await client.list_education_classes(params=params)

    @mcp.tool(
        name="get_education_class",
        description="""get_education_class: GET /education/classes/{id}



Get a specific education class.""",
        tags={"education"},
    )
    async def get_education_class(
        class_id: str = Field(..., description="Education class ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_education_class: GET /education/classes/{id}"""
        client = await get_client()
        return await client.get_education_class(class_id=class_id, params=params)

    @mcp.tool(
        name="list_education_schools",
        description="""list_education_schools: GET /education/schools



List education schools.""",
        tags={"education"},
    )
    async def list_education_schools(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_education_schools: GET /education/schools"""
        client = await get_client()
        return await client.list_education_schools(params=params)

    @mcp.tool(
        name="get_education_school",
        description="""get_education_school: GET /education/schools/{id}



Get a specific education school.""",
        tags={"education"},
    )
    async def get_education_school(
        school_id: str = Field(..., description="Education school ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_education_school: GET /education/schools/{id}"""
        client = await get_client()
        return await client.get_education_school(school_id=school_id, params=params)

    @mcp.tool(
        name="list_education_users",
        description="""list_education_users: GET /education/users



List education users.""",
        tags={"education"},
    )
    async def list_education_users(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_education_users: GET /education/users"""
        client = await get_client()
        return await client.list_education_users(params=params)

    @mcp.tool(
        name="list_education_assignments",
        description="""list_education_assignments: GET /education/classes/{id}/assignments



List assignments for an education class.""",
        tags={"education"},
    )
    async def list_education_assignments(
        class_id: str = Field(..., description="Education class ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_education_assignments: GET /education/classes/{id}/assignments"""
        client = await get_client()
        return await client.list_education_assignments(class_id=class_id, params=params)


def register_agreements_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_agreements",
        description="""list_agreements: GET /agreements



List terms-of-use agreements.""",
        tags={"agreements"},
    )
    async def list_agreements(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_agreements: GET /agreements"""
        client = await get_client()
        return await client.list_agreements(params=params)

    @mcp.tool(
        name="get_agreement",
        description="""get_agreement: GET /agreements/{id}



Get a specific agreement.""",
        tags={"agreements"},
    )
    async def get_agreement(
        agreement_id: str = Field(..., description="Agreement ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_agreement: GET /agreements/{id}"""
        client = await get_client()
        return await client.get_agreement(agreement_id=agreement_id, params=params)

    @mcp.tool(
        name="create_agreement",
        description="""create_agreement: POST /agreements



Create a terms-of-use agreement.""",
        tags={"agreements"},
    )
    async def create_agreement(
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_agreement: POST /agreements"""
        client = await get_client()
        return await client.create_agreement(data=data, params=params)

    @mcp.tool(
        name="delete_agreement",
        description="""delete_agreement: DELETE /agreements/{id}



Delete an agreement.""",
        tags={"agreements"},
    )
    async def delete_agreement(
        agreement_id: str = Field(..., description="Agreement ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_agreement: DELETE /agreements/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete agreement"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_agreement(agreement_id=agreement_id, params=params)


def register_places_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_rooms",
        description="""list_rooms: GET /places/microsoft.graph.room



List conference rooms.""",
        tags={"places"},
    )
    async def list_rooms(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_rooms: GET /places/microsoft.graph.room"""
        client = await get_client()
        return await client.list_rooms(params=params)

    @mcp.tool(
        name="list_room_lists",
        description="""list_room_lists: GET /places/microsoft.graph.roomList



List room lists.""",
        tags={"places"},
    )
    async def list_room_lists(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_room_lists: GET /places/microsoft.graph.roomList"""
        client = await get_client()
        return await client.list_room_lists(params=params)

    @mcp.tool(
        name="get_place",
        description="""get_place: GET /places/{id}



Get a specific place (room or room list).""",
        tags={"places"},
    )
    async def get_place(
        place_id: str = Field(..., description="Place ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_place: GET /places/{id}"""
        client = await get_client()
        return await client.get_place(place_id=place_id, params=params)

    @mcp.tool(
        name="update_place",
        description="""update_place: PATCH /places/{id}



Update a place (room).""",
        tags={"places"},
    )
    async def update_place(
        place_id: str = Field(..., description="Place ID"),
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName, capacity, etc."
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """update_place: PATCH /places/{id}"""
        client = await get_client()
        return await client.update_place(place_id=place_id, data=data, params=params)


def register_print_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_printers",
        description="""list_printers: GET /print/printers



List printers registered in the tenant.""",
        tags={"print"},
    )
    async def list_printers(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_printers: GET /print/printers"""
        client = await get_client()
        return await client.list_printers(params=params)

    @mcp.tool(
        name="get_printer",
        description="""get_printer: GET /print/printers/{id}



Get a specific printer.""",
        tags={"print"},
    )
    async def get_printer(
        printer_id: str = Field(..., description="Printer ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_printer: GET /print/printers/{id}"""
        client = await get_client()
        return await client.get_printer(printer_id=printer_id, params=params)

    @mcp.tool(
        name="list_print_jobs",
        description="""list_print_jobs: GET /print/printers/{id}/jobs



List print jobs for a printer.""",
        tags={"print"},
    )
    async def list_print_jobs(
        printer_id: str = Field(..., description="Printer ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_print_jobs: GET /print/printers/{id}/jobs"""
        client = await get_client()
        return await client.list_print_jobs(printer_id=printer_id, params=params)

    @mcp.tool(
        name="create_print_job",
        description="""create_print_job: POST /print/printers/{id}/jobs



Create a print job.""",
        tags={"print"},
    )
    async def create_print_job(
        printer_id: str = Field(..., description="Printer ID"),
        data: dict[str, Any] | None = Field(None, description="Request body"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_print_job: POST /print/printers/{id}/jobs"""
        client = await get_client()
        return await client.create_print_job(
            printer_id=printer_id, data=data, params=params
        )

    @mcp.tool(
        name="list_print_shares",
        description="""list_print_shares: GET /print/shares



List printer shares.""",
        tags={"print"},
    )
    async def list_print_shares(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_print_shares: GET /print/shares"""
        client = await get_client()
        return await client.list_print_shares(params=params)


def register_privacy_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_subject_rights_requests",
        description="""list_subject_rights_requests: GET /privacy/subjectRightsRequests



List subject rights requests (GDPR/CCPA).""",
        tags={"privacy"},
    )
    async def list_subject_rights_requests(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_subject_rights_requests: GET /privacy/subjectRightsRequests"""
        client = await get_client()
        return await client.list_subject_rights_requests(params=params)

    @mcp.tool(
        name="get_subject_rights_request",
        description="""get_subject_rights_request: GET /privacy/subjectRightsRequests/{id}



Get a specific subject rights request.""",
        tags={"privacy"},
    )
    async def get_subject_rights_request(
        request_id: str = Field(..., description="Subject rights request ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_subject_rights_request: GET /privacy/subjectRightsRequests/{id}"""
        client = await get_client()
        return await client.get_subject_rights_request(
            request_id=request_id, params=params
        )

    @mcp.tool(
        name="create_subject_rights_request",
        description="""create_subject_rights_request: POST /privacy/subjectRightsRequests



Create a subject rights request.""",
        tags={"privacy"},
    )
    async def create_subject_rights_request(
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName, description"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_subject_rights_request: POST /privacy/subjectRightsRequests"""
        client = await get_client()
        return await client.create_subject_rights_request(data=data, params=params)


def register_solutions_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_booking_businesses",
        description="""list_booking_businesses: GET /solutions/bookingBusinesses



List booking businesses.""",
        tags={"solutions"},
    )
    async def list_booking_businesses(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_booking_businesses: GET /solutions/bookingBusinesses"""
        client = await get_client()
        return await client.list_booking_businesses(params=params)

    @mcp.tool(
        name="get_booking_business",
        description="""get_booking_business: GET /solutions/bookingBusinesses/{id}



Get a specific booking business.""",
        tags={"solutions"},
    )
    async def get_booking_business(
        business_id: str = Field(..., description="Booking business ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_booking_business: GET /solutions/bookingBusinesses/{id}"""
        client = await get_client()
        return await client.get_booking_business(business_id=business_id, params=params)

    @mcp.tool(
        name="list_booking_appointments",
        description="""list_booking_appointments: GET /solutions/bookingBusinesses/{id}/appointments



List appointments for a booking business.""",
        tags={"solutions"},
    )
    async def list_booking_appointments(
        business_id: str = Field(..., description="Booking business ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_booking_appointments: GET /solutions/bookingBusinesses/{id}/appointments"""
        client = await get_client()
        return await client.list_booking_appointments(
            business_id=business_id, params=params
        )

    @mcp.tool(
        name="create_booking_appointment",
        description="""create_booking_appointment: POST /solutions/bookingBusinesses/{id}/appointments



Create a booking appointment.""",
        tags={"solutions"},
    )
    async def create_booking_appointment(
        business_id: str = Field(..., description="Booking business ID"),
        data: dict[str, Any] | None = Field(
            None, description="Request body with serviceId, customerName"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_booking_appointment: POST /solutions/bookingBusinesses/{id}/appointments"""
        client = await get_client()
        return await client.create_booking_appointment(
            business_id=business_id, data=data, params=params
        )

    @mcp.tool(
        name="list_virtual_events",
        description="""list_virtual_events: GET /solutions/virtualEvents/townhalls



List virtual event townhalls.""",
        tags={"solutions"},
    )
    async def list_virtual_events(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_virtual_events: GET /solutions/virtualEvents/townhalls"""
        client = await get_client()
        return await client.list_virtual_events(params=params)


def register_storage_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_file_storage_containers",
        description="""list_file_storage_containers: GET /storage/fileStorage/containers



List file storage containers.""",
        tags={"storage"},
    )
    async def list_file_storage_containers(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_file_storage_containers: GET /storage/fileStorage/containers"""
        client = await get_client()
        return await client.list_file_storage_containers(params=params)

    @mcp.tool(
        name="get_file_storage_container",
        description="""get_file_storage_container: GET /storage/fileStorage/containers/{id}



Get a specific file storage container.""",
        tags={"storage"},
    )
    async def get_file_storage_container(
        container_id: str = Field(..., description="File storage container ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_file_storage_container: GET /storage/fileStorage/containers/{id}"""
        client = await get_client()
        return await client.get_file_storage_container(
            container_id=container_id, params=params
        )

    @mcp.tool(
        name="create_file_storage_container",
        description="""create_file_storage_container: POST /storage/fileStorage/containers



Create a file storage container.""",
        tags={"storage"},
    )
    async def create_file_storage_container(
        data: dict[str, Any] | None = Field(
            None, description="Request body with displayName, containerTypeId"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_file_storage_container: POST /storage/fileStorage/containers"""
        client = await get_client()
        return await client.create_file_storage_container(data=data, params=params)


def register_employee_experience_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_learning_providers",
        description="""list_learning_providers: GET /employeeExperience/learningProviders



List learning providers.""",
        tags={"employee_experience"},
    )
    async def list_learning_providers(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_learning_providers: GET /employeeExperience/learningProviders"""
        client = await get_client()
        return await client.list_learning_providers(params=params)

    @mcp.tool(
        name="get_learning_provider",
        description="""get_learning_provider: GET /employeeExperience/learningProviders/{id}



Get a specific learning provider.""",
        tags={"employee_experience"},
    )
    async def get_learning_provider(
        provider_id: str = Field(..., description="Learning provider ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_learning_provider: GET /employeeExperience/learningProviders/{id}"""
        client = await get_client()
        return await client.get_learning_provider(
            provider_id=provider_id, params=params
        )

    @mcp.tool(
        name="list_learning_course_activities",
        description="""list_learning_course_activities: GET /me/employeeExperience/learningCourseActivities



List learning course activities for the current user.""",
        tags={"employee_experience"},
    )
    async def list_learning_course_activities(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_learning_course_activities: GET /me/employeeExperience/learningCourseActivities"""
        client = await get_client()
        return await client.list_learning_course_activities(params=params)


def register_connections_tools(mcp: FastMCP):
    @mcp.tool(
        name="list_external_connections",
        description="""list_external_connections: GET /external/connections



List Microsoft Search external connections.""",
        tags={"connections"},
    )
    async def list_external_connections(
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """list_external_connections: GET /external/connections"""
        client = await get_client()
        return await client.list_external_connections(params=params)

    @mcp.tool(
        name="get_external_connection",
        description="""get_external_connection: GET /external/connections/{id}



Get a specific external connection.""",
        tags={"connections"},
    )
    async def get_external_connection(
        connection_id: str = Field(..., description="External connection ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """get_external_connection: GET /external/connections/{id}"""
        client = await get_client()
        return await client.get_external_connection(
            connection_id=connection_id, params=params
        )

    @mcp.tool(
        name="create_external_connection",
        description="""create_external_connection: POST /external/connections



Create an external connection for Microsoft Search.""",
        tags={"connections"},
    )
    async def create_external_connection(
        data: dict[str, Any] | None = Field(
            None, description="Request body with id, name, description"
        ),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """create_external_connection: POST /external/connections"""
        client = await get_client()
        return await client.create_external_connection(data=data, params=params)

    @mcp.tool(
        name="delete_external_connection",
        description="""delete_external_connection: DELETE /external/connections/{id}



Delete an external connection.""",
        tags={"connections"},
    )
    async def delete_external_connection(
        connection_id: str = Field(..., description="External connection ID"),
        params: dict[str, Any] | None = Field(None, description="Query parameters"),
        ctx: Context = Field(
            description="MCP context for progress reporting", default=None
        ),
    ) -> Any:
        """delete_external_connection: DELETE /external/connections/{id}"""
        if not await ctx_confirm_destructive(ctx, "delete external connection"):
            return {"status": "cancelled", "message": "Operation cancelled by user"}
        await ctx_progress(ctx, 0, 100)
        client = await get_client()
        return await client.delete_external_connection(
            connection_id=connection_id, params=params
        )


def get_mcp_instance() -> tuple[Any, Any, Any, Any]:
    """Initialize and return the MCP instance, args, and middlewares."""
    load_dotenv(find_dotenv())

    args, mcp, middlewares = create_mcp_server(
        name="Microsoft",
        version=__version__,
        instructions="Microsoft Agent MCP Server - Manage emails, calendar, files, and other Microsoft Graph resources.",
    )

    DEFAULT_MISCTOOL = to_boolean(os.getenv("MISCTOOL", "True"))
    if DEFAULT_MISCTOOL:
        register_misc_tools(mcp)
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
    register_prompts(mcp)

    for mw in middlewares:
        mcp.add_middleware(mw)
    registered_tags: list[str] = []
    return mcp, args, middlewares, registered_tags


def mcp_server() -> None:
    mcp, args, middlewares, registered_tags = get_mcp_instance()
    print(f"{'microsoft-agent'} MCP v{__version__}", file=sys.stderr)
    print("Starting MCP Server", file=sys.stderr)
    print(f"  Transport: {args.transport.upper()}", file=sys.stderr)
    print(f"  Auth: {args.auth_type}", file=sys.stderr)
    print(f"  Dynamic Tags Loaded: {len(set(registered_tags))}", file=sys.stderr)

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
