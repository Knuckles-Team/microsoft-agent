#!/usr/bin/python
# coding: utf-8

import os
import argparse
import sys
import logging
from typing import Optional, List, Dict, Union, Any

import requests
from pydantic import Field
from eunomia_mcp.middleware import EunomiaMcpMiddleware
from fastmcp import FastMCP
from fastmcp.server.auth.oidc_proxy import OIDCProxy
from fastmcp.server.auth import OAuthProxy, RemoteAuthProvider
from fastmcp.server.auth.providers.jwt import JWTVerifier, StaticTokenVerifier
from fastmcp.server.middleware.logging import LoggingMiddleware
from fastmcp.server.middleware.timing import TimingMiddleware
from fastmcp.server.middleware.rate_limiting import RateLimitingMiddleware
from fastmcp.server.middleware.error_handling import ErrorHandlingMiddleware
from fastmcp.utilities.logging import get_logger
from microsoft_agent.auth import AuthManager
from microsoft_agent.utils import to_boolean, to_integer
from microsoft_agent.middlewares import (
    UserTokenMiddleware,
    JWTClaimsLoggingMiddleware,
    get_client,
)
from starlette.requests import Request
from starlette.responses import JSONResponse

__version__ = "0.2.3"
print(f"Microsoft MCP v{__version__}")

logger = get_logger(name="TokenMiddleware")
logger.setLevel(logging.DEBUG)

config = {
    "enable_delegation": to_boolean(os.environ.get("ENABLE_DELEGATION", "False")),
    "audience": os.environ.get("AUDIENCE", None),
    "delegated_scopes": os.environ.get("DELEGATED_SCOPES", "api"),
    "token_endpoint": None,  # Will be fetched dynamically from OIDC config
    "oidc_client_id": os.environ.get("OIDC_CLIENT_ID", None),
    "oidc_client_secret": os.environ.get("OIDC_CLIENT_SECRET", None),
    "oidc_config_url": os.environ.get("OIDC_CONFIG_URL", None),
    "jwt_jwks_uri": os.getenv("FASTMCP_SERVER_AUTH_JWT_JWKS_URI", None),
    "jwt_issuer": os.getenv("FASTMCP_SERVER_AUTH_JWT_ISSUER", None),
    "jwt_audience": os.getenv("FASTMCP_SERVER_AUTH_JWT_AUDIENCE", None),
    "jwt_algorithm": os.getenv("FASTMCP_SERVER_AUTH_JWT_ALGORITHM", None),
    "jwt_secret": os.getenv("FASTMCP_SERVER_AUTH_JWT_PUBLIC_KEY", None),
    "jwt_required_scopes": os.getenv("FASTMCP_SERVER_AUTH_JWT_REQUIRED_SCOPES", None),
}

DEFAULT_TRANSPORT = os.getenv("TRANSPORT", "stdio")
DEFAULT_HOST = os.getenv("HOST", "0.0.0.0")
DEFAULT_PORT = to_integer(string=os.getenv("PORT", "8000"))


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

    @mcp.prompt(name="calendar_today", description="Show today's calendar events.")
    def calendar_today() -> str:
        """Show calendar."""
        return "Please show my calendar events for today."


def register_tools(mcp: FastMCP):
    @mcp.custom_route("/health", methods=["GET"])
    async def health_check(request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    # Initialize AuthManager
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
    ]

    _ = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)

    @mcp.tool(
        name="login",
        description="Authenticate with Microsoft using device code flow",
        tags={"auth"},
    )
    def login(
        force: bool = Field(
            False, description="Force a new login even if already logged in"
        )
    ) -> Any:
        """Authenticate with Microsoft using device code flow"""
        client = get_client()
        return client.login(force=force)

    @mcp.tool(
        name="logout", description="Log out from Microsoft account", tags={"auth"}
    )
    def logout() -> Any:
        """Log out from Microsoft account"""
        client = get_client()
        return client.logout()

    @mcp.tool(
        name="verify_login",
        description="Check current Microsoft authentication status",
        tags={"auth"},
    )
    def verify_login() -> Any:
        """Check current Microsoft authentication status"""
        client = get_client()
        return client.verify_login()

    @mcp.tool(
        name="list_accounts",
        description="List all available Microsoft accounts",
        tags={"auth"},
    )
    def list_accounts() -> Any:
        """List all available Microsoft accounts"""
        client = get_client()
        return client.list_accounts()

    @mcp.tool(
        name="search_tools",
        description="Search available Microsoft Graph API tools",
        tags={"meta"},
    )
    def search_tools(
        query: str = Field(..., description="Search query"),
        limit: int = Field(20, description="Max results"),
    ) -> Any:
        """Search available Microsoft Graph API tools"""
        client = get_client()
        return client.search_tools(query=query, limit=limit)

    @mcp.tool(
        name="list_mail_messages",
        description="list_mail_messages: GET /me/messages\n\nTIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:john AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter",
        tags={"mail", "files", "user"},
    )
    def list_mail_messages(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_mail_messages: GET /me/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:john AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return client.list_mail_messages(params=params)

    @mcp.tool(
        name="list_mail_folders",
        description="list_mail_folders: GET /me/mailFolders",
        tags={"mail", "files"},
    )
    def list_mail_folders(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_mail_folders: GET /me/mailFolders"""
        client = get_client()
        return client.list_mail_folders(params=params)

    @mcp.tool(
        name="list_mail_folder_messages",
        description="list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages\n\nTIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter",
        tags={"mail", "files", "user"},
    )
    def list_mail_folder_messages(
        mailFolder_id: str = Field(..., description="Parameter for mailFolder-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return client.list_mail_folder_messages(
            mailFolder_id=mailFolder_id, params=params
        )

    @mcp.tool(
        name="get_mail_message",
        description="get_mail_message: GET /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    def get_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_mail_message: GET /me/messages/{message-id}"""
        client = get_client()
        return client.get_mail_message(message_id=message_id, params=params)

    @mcp.tool(
        name="send_mail",
        description="send_mail: POST /me/sendMail\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"mail"},
    )
    def send_mail(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """send_mail: POST /me/sendMail

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.send_mail(data=data, params=params)

    @mcp.tool(
        name="list_shared_mailbox_messages",
        description="list_shared_mailbox_messages: GET /users/{user-id}/messages\n\nTIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter",
        tags={"mail", "files", "user"},
    )
    def list_shared_mailbox_messages(
        user_id: str = Field(..., description="Parameter for user-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_shared_mailbox_messages: GET /users/{user-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return client.list_shared_mailbox_messages(user_id=user_id, params=params)

    @mcp.tool(
        name="list_shared_mailbox_folder_messages",
        description="list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages\n\nTIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter",
        tags={"mail", "files", "user"},
    )
    def list_shared_mailbox_folder_messages(
        user_id: str = Field(..., description="Parameter for user-id"),
        mailFolder_id: str = Field(..., description="Parameter for mailFolder-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages

        TIP: CRITICAL: When searching emails, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'from:', 'subject:', 'body:', 'to:', 'cc:', 'bcc:', 'attachment:', 'hasAttachments:', 'importance:', 'received:', 'sent:'. Examples: $search='from:john@example.com' | $search='subject:meeting AND hasAttachments:true' | $search='body:urgent AND received>=2024-01-01' | $search='from:alice AND importance:high'. Remember: ALWAYS wrap the entire search expression in double quotes! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return client.list_shared_mailbox_folder_messages(
            user_id=user_id, mailFolder_id=mailFolder_id, params=params
        )

    @mcp.tool(
        name="get_shared_mailbox_message",
        description="get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id}",
        tags={"mail", "user"},
    )
    def get_shared_mailbox_message(
        user_id: str = Field(..., description="Parameter for user-id"),
        message_id: str = Field(..., description="Parameter for message-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id}"""
        client = get_client()
        return client.get_shared_mailbox_message(
            user_id=user_id, message_id=message_id, params=params
        )

    @mcp.tool(
        name="send_shared_mailbox_mail",
        description="send_shared_mailbox_mail: POST /users/{user-id}/sendMail\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"mail"},
    )
    def send_shared_mailbox_mail(
        user_id: str = Field(..., description="Parameter for user-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """send_shared_mailbox_mail: POST /users/{user-id}/sendMail

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.send_shared_mailbox_mail(
            user_id=user_id, data=data, params=params
        )

    @mcp.tool(
        name="list_users",
        description="list_users: GET /users\n\nTIP: CRITICAL: This request requires the ConsistencyLevel header set to eventual. When searching users, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'displayName:'. Examples: $search='displayName:john' | $search='displayName:john' OR 'displayName:jane'. Remember: ALWAYS wrap the entire search expression in double quotes and set the ConsistencyLevel header to eventual! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter",
        tags={"files", "user"},
    )
    def list_users(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_users: GET /users

        TIP: CRITICAL: This request requires the ConsistencyLevel header set to eventual. When searching users, the $search parameter value MUST be wrapped in double quotes. Format: $search='your search query here'. Use KQL (Keyword Query Language) syntax to search specific properties: 'displayName:'. Examples: $search='displayName:john' | $search='displayName:john' OR 'displayName:jane'. Remember: ALWAYS wrap the entire search expression in double quotes and set the ConsistencyLevel header to eventual! Reference: https://learn.microsoft.com/en-us/graph/search-query-parameter
        """
        client = get_client()
        return client.list_users(params=params)

    @mcp.tool(
        name="create_draft_email",
        description="create_draft_email: POST /me/messages",
        tags={"mail"},
    )
    def create_draft_email(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_draft_email: POST /me/messages"""
        client = get_client()
        return client.create_draft_email(data=data, params=params)

    @mcp.tool(
        name="delete_mail_message",
        description="delete_mail_message: DELETE /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    def delete_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_mail_message: DELETE /me/messages/{message-id}"""
        client = get_client()
        return client.delete_mail_message(message_id=message_id, params=params)

    @mcp.tool(
        name="move_mail_message",
        description="move_mail_message: POST /me/messages/{message-id}/move",
        tags={"mail", "user"},
    )
    def move_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """move_mail_message: POST /me/messages/{message-id}/move"""
        client = get_client()
        return client.move_mail_message(message_id=message_id, data=data, params=params)

    @mcp.tool(
        name="update_mail_message",
        description="update_mail_message: PATCH /me/messages/{message-id}",
        tags={"mail", "user"},
    )
    def update_mail_message(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_mail_message: PATCH /me/messages/{message-id}"""
        client = get_client()
        return client.update_mail_message(
            message_id=message_id, data=data, params=params
        )

    @mcp.tool(
        name="add_mail_attachment",
        description="add_mail_attachment: POST /me/messages/{message-id}/attachments",
        tags={"mail", "user"},
    )
    def add_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """add_mail_attachment: POST /me/messages/{message-id}/attachments"""
        client = get_client()
        return client.add_mail_attachment(
            message_id=message_id, data=data, params=params
        )

    @mcp.tool(
        name="list_mail_attachments",
        description="list_mail_attachments: GET /me/messages/{message-id}/attachments",
        tags={"mail", "files", "user"},
    )
    def list_mail_attachments(
        message_id: str = Field(..., description="Parameter for message-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_mail_attachments: GET /me/messages/{message-id}/attachments"""
        client = get_client()
        return client.list_mail_attachments(message_id=message_id, params=params)

    @mcp.tool(
        name="get_mail_attachment",
        description="get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id}",
        tags={"mail", "user"},
    )
    def get_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        attachment_id: str = Field(..., description="Parameter for attachment-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id}"""
        client = get_client()
        return client.get_mail_attachment(
            message_id=message_id, attachment_id=attachment_id, params=params
        )

    @mcp.tool(
        name="delete_mail_attachment",
        description="delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id}",
        tags={"mail", "user"},
    )
    def delete_mail_attachment(
        message_id: str = Field(..., description="Parameter for message-id"),
        attachment_id: str = Field(..., description="Parameter for attachment-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id}"""
        client = get_client()
        return client.delete_mail_attachment(
            message_id=message_id, attachment_id=attachment_id, params=params
        )

    @mcp.tool(
        name="list_calendar_events",
        description="list_calendar_events: GET /me/events",
        tags={"calendar", "files"},
    )
    def list_calendar_events(
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
        timezone: Optional[str] = Field(None, description="IANA timezone"),
    ) -> Any:
        """list_calendar_events: GET /me/events"""
        client = get_client()
        return client.list_calendar_events(params=params, timezone=timezone)

    @mcp.tool(
        name="get_calendar_event",
        description="get_calendar_event: GET /me/events/{event-id}",
        tags={"calendar"},
    )
    def get_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
        timezone: Optional[str] = Field(None, description="IANA timezone"),
    ) -> Any:
        """get_calendar_event: GET /me/events/{event-id}"""
        client = get_client()
        return client.get_calendar_event(
            event_id=event_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="create_calendar_event",
        description="create_calendar_event: POST /me/events\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"calendar"},
    )
    def create_calendar_event(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_calendar_event: POST /me/events

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.create_calendar_event(data=data, params=params)

    @mcp.tool(
        name="update_calendar_event",
        description="update_calendar_event: PATCH /me/events/{event-id}\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"calendar"},
    )
    def update_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_calendar_event: PATCH /me/events/{event-id}

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.update_calendar_event(event_id=event_id, data=data, params=params)

    @mcp.tool(
        name="delete_calendar_event",
        description="delete_calendar_event: DELETE /me/events/{event-id}",
        tags={"calendar"},
    )
    def delete_calendar_event(
        event_id: str = Field(..., description="Parameter for event-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_calendar_event: DELETE /me/events/{event-id}"""
        client = get_client()
        return client.delete_calendar_event(event_id=event_id, params=params)

    @mcp.tool(
        name="list_specific_calendar_events",
        description="list_specific_calendar_events: GET /me/calendars/{calendar-id}/events",
        tags={"calendar", "files"},
    )
    def list_specific_calendar_events(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
        timezone: Optional[str] = Field(None, description="IANA timezone"),
    ) -> Any:
        """list_specific_calendar_events: GET /me/calendars/{calendar-id}/events"""
        client = get_client()
        return client.list_specific_calendar_events(
            calendar_id=calendar_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="get_specific_calendar_event",
        description="get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id}",
        tags={"calendar"},
    )
    def get_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
        timezone: Optional[str] = Field(None, description="IANA timezone"),
    ) -> Any:
        """get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id}"""
        client = get_client()
        return client.get_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, params=params, timezone=timezone
        )

    @mcp.tool(
        name="create_specific_calendar_event",
        description="create_specific_calendar_event: POST /me/calendars/{calendar-id}/events\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"calendar"},
    )
    def create_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_specific_calendar_event: POST /me/calendars/{calendar-id}/events

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.create_specific_calendar_event(
            calendar_id=calendar_id, data=data, params=params
        )

    @mcp.tool(
        name="update_specific_calendar_event",
        description="update_specific_calendar_event: PATCH /me/calendars/{calendar-id}/events/{event-id}\n\nTIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.",
        tags={"calendar"},
    )
    def update_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_specific_calendar_event: PATCH /me/calendars/{calendar-id}/events/{event-id}

        TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
        """
        client = get_client()
        return client.update_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_specific_calendar_event",
        description="delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id}",
        tags={"calendar"},
    )
    def delete_specific_calendar_event(
        calendar_id: str = Field(..., description="Parameter for calendar-id"),
        event_id: str = Field(..., description="Parameter for event-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id}"""
        client = get_client()
        return client.delete_specific_calendar_event(
            calendar_id=calendar_id, event_id=event_id, params=params
        )

    @mcp.tool(
        name="get_calendar_view",
        description="get_calendar_view: GET /me/calendarView",
        tags={"calendar"},
    )
    def get_calendar_view(
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
        timezone: Optional[str] = Field(None, description="IANA timezone"),
    ) -> Any:
        """get_calendar_view: GET /me/calendarView"""
        client = get_client()
        return client.get_calendar_view(params=params, timezone=timezone)

    @mcp.tool(
        name="list_calendars",
        description="list_calendars: GET /me/calendars",
        tags={"calendar", "files"},
    )
    def list_calendars(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_calendars: GET /me/calendars"""
        client = get_client()
        return client.list_calendars(params=params)

    @mcp.tool(
        name="find_meeting_times",
        description="find_meeting_times: POST /me/findMeetingTimes",
        tags={"calendar", "user"},
    )
    def find_meeting_times(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """find_meeting_times: POST /me/findMeetingTimes"""
        client = get_client()
        return client.find_meeting_times(data=data, params=params)

    @mcp.tool(
        name="list_drives", description="list_drives: GET /me/drives", tags={"files"}
    )
    def list_drives(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_drives: GET /me/drives"""
        client = get_client()
        return client.list_drives(params=params)

    @mcp.tool(
        name="get_drive_root_item",
        description="get_drive_root_item: GET /drives/{drive-id}/root",
        tags={"files"},
    )
    def get_drive_root_item(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_drive_root_item: GET /drives/{drive-id}/root"""
        client = get_client()
        return client.get_drive_root_item(drive_id=drive_id, params=params)

    @mcp.tool(
        name="get_root_folder",
        description="get_root_folder: GET /drives/{drive-id}/root",
        tags={"mail", "files"},
    )
    def get_root_folder(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_root_folder: GET /drives/{drive-id}/root"""
        client = get_client()
        return client.get_root_folder(drive_id=drive_id, params=params)

    @mcp.tool(
        name="list_folder_files",
        description="list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children",
        tags={"mail", "files"},
    )
    def list_folder_files(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children"""
        client = get_client()
        return client.list_folder_files(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="download_onedrive_file_content",
        description="download_onedrive_file_content: GET /drives/{drive-id}/items/{driveItem-id}/content",
        tags={"files"},
    )
    def download_onedrive_file_content(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """download_onedrive_file_content: GET /drives/{drive-id}/items/{driveItem-id}/content"""
        client = get_client()
        return client.download_onedrive_file_content(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="delete_onedrive_file",
        description="delete_onedrive_file: DELETE /drives/{drive-id}/items/{driveItem-id}",
        tags={"files"},
    )
    def delete_onedrive_file(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_onedrive_file: DELETE /drives/{drive-id}/items/{driveItem-id}"""
        client = get_client()
        return client.delete_onedrive_file(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="upload_file_content",
        description="upload_file_content: PUT /drives/{drive-id}/items/{driveItem-id}/content",
        tags={"files"},
    )
    def upload_file_content(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """upload_file_content: PUT /drives/{drive-id}/items/{driveItem-id}/content"""
        client = get_client()
        return client.upload_file_content(
            drive_id=drive_id, driveItem_id=driveItem_id, data=data, params=params
        )

    @mcp.tool(
        name="create_excel_chart",
        description="create_excel_chart: POST /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add",
        tags={"files"},
    )
    def create_excel_chart(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_excel_chart: POST /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add"""
        client = get_client()
        return client.create_excel_chart(
            drive_id=drive_id,
            driveItem_id=driveItem_id,
            workbookWorksheet_id=workbookWorksheet_id,
            data=data,
            params=params,
        )

    @mcp.tool(
        name="format_excel_range",
        description="format_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/format",
        tags={"files"},
    )
    def format_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """format_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/format"""
        client = get_client()
        return client.format_excel_range(
            drive_id=drive_id,
            driveItem_id=driveItem_id,
            workbookWorksheet_id=workbookWorksheet_id,
            data=data,
            params=params,
        )

    @mcp.tool(
        name="sort_excel_range",
        description="sort_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/sort",
        tags={"files"},
    )
    def sort_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """sort_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/sort"""
        client = get_client()
        return client.sort_excel_range(
            drive_id=drive_id,
            driveItem_id=driveItem_id,
            workbookWorksheet_id=workbookWorksheet_id,
            data=data,
            params=params,
        )

    @mcp.tool(
        name="get_excel_range",
        description="get_excel_range: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')",
        tags={"files"},
    )
    def get_excel_range(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        workbookWorksheet_id: str = Field(
            ..., description="Parameter for workbookWorksheet-id"
        ),
        address: str = Field(..., description="Parameter for address"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_excel_range: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range(address='{address}')"""
        client = get_client()
        return client.get_excel_range(
            drive_id=drive_id,
            driveItem_id=driveItem_id,
            workbookWorksheet_id=workbookWorksheet_id,
            address=address,
            params=params,
        )

    @mcp.tool(
        name="list_excel_worksheets",
        description="list_excel_worksheets: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets",
        tags={"files"},
    )
    def list_excel_worksheets(
        drive_id: str = Field(..., description="Parameter for drive-id"),
        driveItem_id: str = Field(..., description="Parameter for driveItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_excel_worksheets: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets"""
        client = get_client()
        return client.list_excel_worksheets(
            drive_id=drive_id, driveItem_id=driveItem_id, params=params
        )

    @mcp.tool(
        name="list_onenote_notebooks",
        description="list_onenote_notebooks: GET /me/onenote/notebooks",
        tags={"files", "notes"},
    )
    def list_onenote_notebooks(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_onenote_notebooks: GET /me/onenote/notebooks"""
        client = get_client()
        return client.list_onenote_notebooks(params=params)

    @mcp.tool(
        name="list_onenote_notebook_sections",
        description="list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections",
        tags={"files", "notes"},
    )
    def list_onenote_notebook_sections(
        notebook_id: str = Field(..., description="Parameter for notebook-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections"""
        client = get_client()
        return client.list_onenote_notebook_sections(
            notebook_id=notebook_id, params=params
        )

    @mcp.tool(
        name="list_onenote_section_pages",
        description="list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages",
        tags={"files", "notes"},
    )
    def list_onenote_section_pages(
        onenoteSection_id: str = Field(
            ..., description="Parameter for onenoteSection-id"
        ),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages"""
        client = get_client()
        return client.list_onenote_section_pages(
            onenoteSection_id=onenoteSection_id, params=params
        )

    @mcp.tool(
        name="get_onenote_page_content",
        description="get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content",
        tags={"notes"},
    )
    def get_onenote_page_content(
        onenotePage_id: str = Field(..., description="Parameter for onenotePage-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content"""
        client = get_client()
        return client.get_onenote_page_content(
            onenotePage_id=onenotePage_id, params=params
        )

    @mcp.tool(
        name="create_onenote_page",
        description="create_onenote_page: POST /me/onenote/pages",
        tags={"notes"},
    )
    def create_onenote_page(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_onenote_page: POST /me/onenote/pages"""
        client = get_client()
        return client.create_onenote_page(data=data, params=params)

    @mcp.tool(
        name="list_todo_task_lists",
        description="list_todo_task_lists: GET /me/todo/lists",
        tags={"files", "tasks"},
    )
    def list_todo_task_lists(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_todo_task_lists: GET /me/todo/lists"""
        client = get_client()
        return client.list_todo_task_lists(params=params)

    @mcp.tool(
        name="list_todo_tasks",
        description="list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks",
        tags={"files", "tasks"},
    )
    def list_todo_tasks(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks"""
        client = get_client()
        return client.list_todo_tasks(todoTaskList_id=todoTaskList_id, params=params)

    @mcp.tool(
        name="get_todo_task",
        description="get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}",
        tags={"tasks"},
    )
    def get_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        client = get_client()
        return client.get_todo_task(
            todoTaskList_id=todoTaskList_id, todoTask_id=todoTask_id, params=params
        )

    @mcp.tool(
        name="create_todo_task",
        description="create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks",
        tags={"tasks"},
    )
    def create_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks"""
        client = get_client()
        return client.create_todo_task(
            todoTaskList_id=todoTaskList_id, data=data, params=params
        )

    @mcp.tool(
        name="update_todo_task",
        description="update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}",
        tags={"tasks"},
    )
    def update_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        client = get_client()
        return client.update_todo_task(
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
    def delete_todo_task(
        todoTaskList_id: str = Field(..., description="Parameter for todoTaskList-id"),
        todoTask_id: str = Field(..., description="Parameter for todoTask-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_todo_task: DELETE /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}"""
        client = get_client()
        return client.delete_todo_task(
            todoTaskList_id=todoTaskList_id, todoTask_id=todoTask_id, params=params
        )

    @mcp.tool(
        name="list_planner_tasks",
        description="list_planner_tasks: GET /me/planner/tasks",
        tags={"files", "tasks"},
    )
    def list_planner_tasks(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_planner_tasks: GET /me/planner/tasks"""
        client = get_client()
        return client.list_planner_tasks(params=params)

    @mcp.tool(
        name="get_planner_plan",
        description="get_planner_plan: GET /planner/plans/{plannerPlan-id}",
        tags={"tasks"},
    )
    def get_planner_plan(
        plannerPlan_id: str = Field(..., description="Parameter for plannerPlan-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_planner_plan: GET /planner/plans/{plannerPlan-id}"""
        client = get_client()
        return client.get_planner_plan(plannerPlan_id=plannerPlan_id, params=params)

    @mcp.tool(
        name="list_plan_tasks",
        description="list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks",
        tags={"files", "tasks"},
    )
    def list_plan_tasks(
        plannerPlan_id: str = Field(..., description="Parameter for plannerPlan-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks"""
        client = get_client()
        return client.list_plan_tasks(plannerPlan_id=plannerPlan_id, params=params)

    @mcp.tool(
        name="get_planner_task",
        description="get_planner_task: GET /planner/tasks/{plannerTask-id}",
        tags={"tasks"},
    )
    def get_planner_task(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_planner_task: GET /planner/tasks/{plannerTask-id}"""
        client = get_client()
        return client.get_planner_task(plannerTask_id=plannerTask_id, params=params)

    @mcp.tool(
        name="create_planner_task",
        description="create_planner_task: POST /planner/tasks",
        tags={"tasks"},
    )
    def create_planner_task(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_planner_task: POST /planner/tasks"""
        client = get_client()
        return client.create_planner_task(data=data, params=params)

    @mcp.tool(
        name="update_planner_task",
        description="update_planner_task: PATCH /planner/tasks/{plannerTask-id}",
        tags={"tasks"},
    )
    def update_planner_task(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_planner_task: PATCH /planner/tasks/{plannerTask-id}"""
        client = get_client()
        return client.update_planner_task(
            plannerTask_id=plannerTask_id, data=data, params=params
        )

    @mcp.tool(
        name="update_planner_task_details",
        description="update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details",
        tags={"tasks"},
    )
    def update_planner_task_details(
        plannerTask_id: str = Field(..., description="Parameter for plannerTask-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details"""
        client = get_client()
        return client.update_planner_task_details(
            plannerTask_id=plannerTask_id, data=data, params=params
        )

    @mcp.tool(
        name="list_outlook_contacts",
        description="list_outlook_contacts: GET /me/contacts",
        tags={"files", "contacts"},
    )
    def list_outlook_contacts(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_outlook_contacts: GET /me/contacts"""
        client = get_client()
        return client.list_outlook_contacts(params=params)

    @mcp.tool(
        name="get_outlook_contact",
        description="get_outlook_contact: GET /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    def get_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_outlook_contact: GET /me/contacts/{contact-id}"""
        client = get_client()
        return client.get_outlook_contact(contact_id=contact_id, params=params)

    @mcp.tool(
        name="create_outlook_contact",
        description="create_outlook_contact: POST /me/contacts",
        tags={"contacts"},
    )
    def create_outlook_contact(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """create_outlook_contact: POST /me/contacts"""
        client = get_client()
        return client.create_outlook_contact(data=data, params=params)

    @mcp.tool(
        name="update_outlook_contact",
        description="update_outlook_contact: PATCH /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    def update_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """update_outlook_contact: PATCH /me/contacts/{contact-id}"""
        client = get_client()
        return client.update_outlook_contact(
            contact_id=contact_id, data=data, params=params
        )

    @mcp.tool(
        name="delete_outlook_contact",
        description="delete_outlook_contact: DELETE /me/contacts/{contact-id}",
        tags={"contacts"},
    )
    def delete_outlook_contact(
        contact_id: str = Field(..., description="Parameter for contact-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """delete_outlook_contact: DELETE /me/contacts/{contact-id}"""
        client = get_client()
        return client.delete_outlook_contact(contact_id=contact_id, params=params)

    @mcp.tool(
        name="get_current_user", description="get_current_user: GET /me", tags={"user"}
    )
    def get_current_user(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """get_current_user: GET /me"""
        client = get_client()
        return client.get_current_user(params=params)

    @mcp.tool(
        name="list_chats",
        description="list_chats: GET /me/chats",
        tags={"files", "chat"},
    )
    def list_chats(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_chats: GET /me/chats"""
        client = get_client()
        return client.list_chats(params=params)

    @mcp.tool(
        name="get_chat", description="get_chat: GET /chats/{chat-id}", tags={"chat"}
    )
    def get_chat(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_chat: GET /chats/{chat-id}"""
        client = get_client()
        return client.get_chat(chat_id=chat_id, params=params)

    @mcp.tool(
        name="list_chat_messages",
        description="list_chat_messages: GET /chats/{chat-id}/messages",
        tags={"mail", "files", "user", "chat"},
    )
    def list_chat_messages(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_chat_messages: GET /chats/{chat-id}/messages"""
        client = get_client()
        return client.list_chat_messages(chat_id=chat_id, params=params)

    @mcp.tool(
        name="get_chat_message",
        description="get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id}",
        tags={"mail", "user", "chat"},
    )
    def get_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id}"""
        client = get_client()
        return client.get_chat_message(
            chat_id=chat_id, chatMessage_id=chatMessage_id, params=params
        )

    @mcp.tool(
        name="send_chat_message",
        description="send_chat_message: POST /chats/{chat-id}/messages",
        tags={"mail", "user", "chat"},
    )
    def send_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """send_chat_message: POST /chats/{chat-id}/messages"""
        client = get_client()
        return client.send_chat_message(chat_id=chat_id, data=data, params=params)

    @mcp.tool(
        name="list_joined_teams",
        description="list_joined_teams: GET /me/joinedTeams",
        tags={"files", "teams"},
    )
    def list_joined_teams(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """list_joined_teams: GET /me/joinedTeams"""
        client = get_client()
        return client.list_joined_teams(params=params)

    @mcp.tool(
        name="get_team", description="get_team: GET /teams/{team-id}", tags={"teams"}
    )
    def get_team(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_team: GET /teams/{team-id}"""
        client = get_client()
        return client.get_team(team_id=team_id, params=params)

    @mcp.tool(
        name="list_team_channels",
        description="list_team_channels: GET /teams/{team-id}/channels",
        tags={"files", "teams"},
    )
    def list_team_channels(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_team_channels: GET /teams/{team-id}/channels"""
        client = get_client()
        return client.list_team_channels(team_id=team_id, params=params)

    @mcp.tool(
        name="get_team_channel",
        description="get_team_channel: GET /teams/{team-id}/channels/{channel-id}",
        tags={"teams"},
    )
    def get_team_channel(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_team_channel: GET /teams/{team-id}/channels/{channel-id}"""
        client = get_client()
        return client.get_team_channel(
            team_id=team_id, channel_id=channel_id, params=params
        )

    @mcp.tool(
        name="list_channel_messages",
        description="list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages",
        tags={"mail", "files", "user", "teams"},
    )
    def list_channel_messages(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages"""
        client = get_client()
        return client.list_channel_messages(
            team_id=team_id, channel_id=channel_id, params=params
        )

    @mcp.tool(
        name="get_channel_message",
        description="get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}",
        tags={"mail", "user", "teams"},
    )
    def get_channel_message(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}"""
        client = get_client()
        return client.get_channel_message(
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
    def send_channel_message(
        team_id: str = Field(..., description="Parameter for team-id"),
        channel_id: str = Field(..., description="Parameter for channel-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages"""
        client = get_client()
        return client.send_channel_message(
            team_id=team_id, channel_id=channel_id, data=data, params=params
        )

    @mcp.tool(
        name="list_team_members",
        description="list_team_members: GET /teams/{team-id}/members",
        tags={"files", "user", "teams"},
    )
    def list_team_members(
        team_id: str = Field(..., description="Parameter for team-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_team_members: GET /teams/{team-id}/members"""
        client = get_client()
        return client.list_team_members(team_id=team_id, params=params)

    @mcp.tool(
        name="list_chat_message_replies",
        description="list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies",
        tags={"mail", "files", "user", "chat"},
    )
    def list_chat_message_replies(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies"""
        client = get_client()
        return client.list_chat_message_replies(
            chat_id=chat_id, chatMessage_id=chatMessage_id, params=params
        )

    @mcp.tool(
        name="reply_to_chat_message",
        description="reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies",
        tags={"mail", "user", "chat"},
    )
    def reply_to_chat_message(
        chat_id: str = Field(..., description="Parameter for chat-id"),
        chatMessage_id: str = Field(..., description="Parameter for chatMessage-id"),
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies"""
        client = get_client()
        return client.reply_to_chat_message(
            chat_id=chat_id, chatMessage_id=chatMessage_id, data=data, params=params
        )

    @mcp.tool(
        name="search_sharepoint_sites",
        description="search_sharepoint_sites: GET /sites",
        tags={"search", "sites"},
    )
    def search_sharepoint_sites(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """search_sharepoint_sites: GET /sites"""
        client = get_client()
        return client.search_sharepoint_sites(params=params)

    @mcp.tool(
        name="get_sharepoint_site",
        description="get_sharepoint_site: GET /sites/{site-id}",
        tags={"sites"},
    )
    def get_sharepoint_site(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site: GET /sites/{site-id}"""
        client = get_client()
        return client.get_sharepoint_site(site_id=site_id, params=params)

    @mcp.tool(
        name="list_sharepoint_site_drives",
        description="list_sharepoint_site_drives: GET /sites/{site-id}/drives",
        tags={"files", "sites"},
    )
    def list_sharepoint_site_drives(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_sharepoint_site_drives: GET /sites/{site-id}/drives"""
        client = get_client()
        return client.list_sharepoint_site_drives(site_id=site_id, params=params)

    @mcp.tool(
        name="get_sharepoint_site_drive_by_id",
        description="get_sharepoint_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id}",
        tags={"files", "sites"},
    )
    def get_sharepoint_site_drive_by_id(
        site_id: str = Field(..., description="Parameter for site-id"),
        drive_id: str = Field(..., description="Parameter for drive-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id}"""
        client = get_client()
        return client.get_sharepoint_site_drive_by_id(
            site_id=site_id, drive_id=drive_id, params=params
        )

    @mcp.tool(
        name="list_sharepoint_site_items",
        description="list_sharepoint_site_items: GET /sites/{site-id}/items",
        tags={"files", "sites"},
    )
    def list_sharepoint_site_items(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_sharepoint_site_items: GET /sites/{site-id}/items"""
        client = get_client()
        return client.list_sharepoint_site_items(site_id=site_id, params=params)

    @mcp.tool(
        name="get_sharepoint_site_item",
        description="get_sharepoint_site_item: GET /sites/{site-id}/items/{baseItem-id}",
        tags={"files", "sites"},
    )
    def get_sharepoint_site_item(
        site_id: str = Field(..., description="Parameter for site-id"),
        baseItem_id: str = Field(..., description="Parameter for baseItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site_item: GET /sites/{site-id}/items/{baseItem-id}"""
        client = get_client()
        return client.get_sharepoint_site_item(
            site_id=site_id, baseItem_id=baseItem_id, params=params
        )

    @mcp.tool(
        name="list_sharepoint_site_lists",
        description="list_sharepoint_site_lists: GET /sites/{site-id}/lists",
        tags={"files", "sites"},
    )
    def list_sharepoint_site_lists(
        site_id: str = Field(..., description="Parameter for site-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_sharepoint_site_lists: GET /sites/{site-id}/lists"""
        client = get_client()
        return client.list_sharepoint_site_lists(site_id=site_id, params=params)

    @mcp.tool(
        name="get_sharepoint_site_list",
        description="get_sharepoint_site_list: GET /sites/{site-id}/lists/{list-id}",
        tags={"files", "sites"},
    )
    def get_sharepoint_site_list(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site_list: GET /sites/{site-id}/lists/{list-id}"""
        client = get_client()
        return client.get_sharepoint_site_list(
            site_id=site_id, list_id=list_id, params=params
        )

    @mcp.tool(
        name="list_sharepoint_site_list_items",
        description="list_sharepoint_site_list_items: GET /sites/{site-id}/lists/{list-id}/items",
        tags={"files", "sites"},
    )
    def list_sharepoint_site_list_items(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """list_sharepoint_site_list_items: GET /sites/{site-id}/lists/{list-id}/items"""
        client = get_client()
        return client.list_sharepoint_site_list_items(
            site_id=site_id, list_id=list_id, params=params
        )

    @mcp.tool(
        name="get_sharepoint_site_list_item",
        description="get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id}",
        tags={"files", "sites"},
    )
    def get_sharepoint_site_list_item(
        site_id: str = Field(..., description="Parameter for site-id"),
        list_id: str = Field(..., description="Parameter for list-id"),
        listItem_id: str = Field(..., description="Parameter for listItem-id"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id}"""
        client = get_client()
        return client.get_sharepoint_site_list_item(
            site_id=site_id, list_id=list_id, listItem_id=listItem_id, params=params
        )

    @mcp.tool(
        name="get_sharepoint_site_by_path",
        description="get_sharepoint_site_by_path: GET /sites/{site-id}/getByPath(path='{path}')",
        tags={"sites"},
    )
    def get_sharepoint_site_by_path(
        site_id: str = Field(..., description="Parameter for site-id"),
        path: str = Field(..., description="Parameter for path"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """get_sharepoint_site_by_path: GET /sites/{site-id}/getByPath(path='{path}')"""
        client = get_client()
        return client.get_sharepoint_site_by_path(
            site_id=site_id, path=path, params=params
        )

    @mcp.tool(
        name="get_sharepoint_sites_delta",
        description="get_sharepoint_sites_delta: GET /sites/delta()",
        tags={"sites"},
    )
    def get_sharepoint_sites_delta(
        params: Optional[Dict[(str, Any)]] = Field(None, description="Query parameters")
    ) -> Any:
        """get_sharepoint_sites_delta: GET /sites/delta()"""
        client = get_client()
        return client.get_sharepoint_sites_delta(params=params)

    @mcp.tool(
        name="search_query",
        description="search_query: POST /search/query",
        tags={"search"},
    )
    def search_query(
        data: Optional[Dict[(str, Any)]] = Field(None, description="Request body data"),
        params: Optional[Dict[(str, Any)]] = Field(
            None, description="Query parameters"
        ),
    ) -> Any:
        """search_query: POST /search/query"""
        client = get_client()
        return client.search_query(data=data, params=params)


def microsoft_mcp() -> None:
    """Run the Microsoft MCP server with specified transport and connection parameters.

    This function parses command-line arguments to configure and start the MCP server for Microsoft API interactions.
    It supports stdio or TCP transport modes and exits on invalid arguments or help requests.

    """
    parser = argparse.ArgumentParser(add_help=False, description="Microsoft MCP Server")
    parser.add_argument(
        "-t",
        "--transport",
        default=DEFAULT_TRANSPORT,
        choices=["stdio", "streamable-http", "sse"],
        help="Transport method: 'stdio', 'streamable-http', or 'sse' [legacy] (default: stdio)",
    )
    parser.add_argument(
        "-s",
        "--host",
        default=DEFAULT_HOST,
        help="Host address for HTTP transport (default: 0.0.0.0)",
    )
    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=DEFAULT_PORT,
        help="Port number for HTTP transport (default: 8000)",
    )
    parser.add_argument(
        "--auth-type",
        default="none",
        choices=["none", "static", "jwt", "oauth-proxy", "oidc-proxy", "remote-oauth"],
        help="Authentication type for MCP server: 'none' (disabled), 'static' (internal), 'jwt' (external token verification), 'oauth-proxy', 'oidc-proxy', 'remote-oauth' (external) (default: none)",
    )
    # JWT/Token params
    parser.add_argument(
        "--token-jwks-uri", default=None, help="JWKS URI for JWT verification"
    )
    parser.add_argument(
        "--token-issuer", default=None, help="Issuer for JWT verification"
    )
    parser.add_argument(
        "--token-audience", default=None, help="Audience for JWT verification"
    )
    parser.add_argument(
        "--token-algorithm",
        default=os.getenv("FASTMCP_SERVER_AUTH_JWT_ALGORITHM"),
        choices=[
            "HS256",
            "HS384",
            "HS512",
            "RS256",
            "RS384",
            "RS512",
            "ES256",
            "ES384",
            "ES512",
        ],
        help="JWT signing algorithm (required for HMAC or static key). Auto-detected for JWKS.",
    )
    parser.add_argument(
        "--token-secret",
        default=os.getenv("FASTMCP_SERVER_AUTH_JWT_PUBLIC_KEY"),
        help="Shared secret for HMAC (HS*) or PEM public key for static asymmetric verification.",
    )
    parser.add_argument(
        "--token-public-key",
        default=os.getenv("FASTMCP_SERVER_AUTH_JWT_PUBLIC_KEY"),
        help="Path to PEM public key file or inline PEM string (for static asymmetric keys).",
    )
    parser.add_argument(
        "--required-scopes",
        default=os.getenv("FASTMCP_SERVER_AUTH_JWT_REQUIRED_SCOPES"),
        help="Comma-separated list of required scopes (e.g., microsoft.read,microsoft.write).",
    )
    # OAuth Proxy params
    parser.add_argument(
        "--oauth-upstream-auth-endpoint",
        default=None,
        help="Upstream authorization endpoint for OAuth Proxy",
    )
    parser.add_argument(
        "--oauth-upstream-token-endpoint",
        default=None,
        help="Upstream token endpoint for OAuth Proxy",
    )
    parser.add_argument(
        "--oauth-upstream-client-id",
        default=None,
        help="Upstream client ID for OAuth Proxy",
    )
    parser.add_argument(
        "--oauth-upstream-client-secret",
        default=None,
        help="Upstream client secret for OAuth Proxy",
    )
    parser.add_argument(
        "--oauth-base-url", default=None, help="Base URL for OAuth Proxy"
    )
    # OIDC Proxy params
    parser.add_argument(
        "--oidc-config-url", default=None, help="OIDC configuration URL"
    )
    parser.add_argument("--oidc-client-id", default=None, help="OIDC client ID")
    parser.add_argument("--oidc-client-secret", default=None, help="OIDC client secret")
    parser.add_argument("--oidc-base-url", default=None, help="Base URL for OIDC Proxy")
    # Remote OAuth params
    parser.add_argument(
        "--remote-auth-servers",
        default=None,
        help="Comma-separated list of authorization servers for Remote OAuth",
    )
    parser.add_argument(
        "--remote-base-url", default=None, help="Base URL for Remote OAuth"
    )
    # Common
    parser.add_argument(
        "--allowed-client-redirect-uris",
        default=None,
        help="Comma-separated list of allowed client redirect URIs",
    )
    # Eunomia params
    parser.add_argument(
        "--eunomia-type",
        default="none",
        choices=["none", "embedded", "remote"],
        help="Eunomia authorization type: 'none' (disabled), 'embedded' (built-in), 'remote' (external) (default: none)",
    )
    parser.add_argument(
        "--eunomia-policy-file",
        default="mcp_policies.json",
        help="Policy file for embedded Eunomia (default: mcp_policies.json)",
    )
    parser.add_argument(
        "--eunomia-remote-url", default=None, help="URL for remote Eunomia server"
    )
    # Delegation params
    parser.add_argument(
        "--enable-delegation",
        action="store_true",
        default=to_boolean(os.environ.get("ENABLE_DELEGATION", "False")),
        help="Enable OIDC token delegation",
    )
    parser.add_argument(
        "--audience",
        default=os.environ.get("AUDIENCE", None),
        help="Audience for the delegated token",
    )
    parser.add_argument(
        "--delegated-scopes",
        default=os.environ.get("DELEGATED_SCOPES", "api"),
        help="Scopes for the delegated token (space-separated)",
    )
    parser.add_argument(
        "--openapi-file",
        default=None,
        help="Path to the OpenAPI JSON file to import additional tools from",
    )
    parser.add_argument(
        "--openapi-base-url",
        default=None,
        help="Base URL for the OpenAPI client (overrides instance URL)",
    )
    parser.add_argument(
        "--openapi-use-token",
        action="store_true",
        help="Use the incoming Bearer token (from MCP request) to authenticate OpenAPI import",
    )

    parser.add_argument(
        "--openapi-username",
        default=os.getenv("OPENAPI_USERNAME"),
        help="Username for basic auth during OpenAPI import",
    )

    parser.add_argument(
        "--openapi-password",
        default=os.getenv("OPENAPI_PASSWORD"),
        help="Password for basic auth during OpenAPI import",
    )

    parser.add_argument(
        "--openapi-client-id",
        default=os.getenv("OPENAPI_CLIENT_ID"),
        help="OAuth client ID for OpenAPI import",
    )

    parser.add_argument(
        "--openapi-client-secret",
        default=os.getenv("OPENAPI_CLIENT_SECRET"),
        help="OAuth client secret for OpenAPI import",
    )

    parser.add_argument("--help", action="store_true", help="Show usage")

    args = parser.parse_args()

    if hasattr(args, "help") and args.help:

        usage()

        sys.exit(0)

    if args.port < 0 or args.port > 65535:
        print(f"Error: Port {args.port} is out of valid range (0-65535).")
        sys.exit(1)

    # Update config with CLI arguments
    config["enable_delegation"] = args.enable_delegation
    config["audience"] = args.audience or config["audience"]
    config["delegated_scopes"] = args.delegated_scopes or config["delegated_scopes"]
    config["oidc_config_url"] = args.oidc_config_url or config["oidc_config_url"]
    config["oidc_client_id"] = args.oidc_client_id or config["oidc_client_id"]
    config["oidc_client_secret"] = (
        args.oidc_client_secret or config["oidc_client_secret"]
    )

    # Configure delegation if enabled
    if config["enable_delegation"]:
        if args.auth_type != "oidc-proxy":
            logger.error("Token delegation requires auth-type=oidc-proxy")
            sys.exit(1)
        if not config["audience"]:
            logger.error("audience is required for delegation")
            sys.exit(1)
        if not all(
            [
                config["oidc_config_url"],
                config["oidc_client_id"],
                config["oidc_client_secret"],
            ]
        ):
            logger.error(
                "Delegation requires complete OIDC configuration (oidc-config-url, oidc-client-id, oidc-client-secret)"
            )
            sys.exit(1)

        # Fetch OIDC configuration to get token_endpoint
        try:
            logger.info(
                "Fetching OIDC configuration",
                extra={"oidc_config_url": config["oidc_config_url"]},
            )
            oidc_config_resp = requests.get(config["oidc_config_url"])
            oidc_config_resp.raise_for_status()
            oidc_config = oidc_config_resp.json()
            config["token_endpoint"] = oidc_config.get("token_endpoint")
            if not config["token_endpoint"]:
                logger.error("No token_endpoint found in OIDC configuration")
                raise ValueError("No token_endpoint found in OIDC configuration")
            logger.info(
                "OIDC configuration fetched successfully",
                extra={"token_endpoint": config["token_endpoint"]},
            )
        except Exception as e:
            print(f"Failed to fetch OIDC configuration: {e}")
            logger.error(
                "Failed to fetch OIDC configuration",
                extra={"error_type": type(e).__name__, "error_message": str(e)},
            )
            sys.exit(1)

    # Set auth based on type
    auth = None
    allowed_uris = (
        args.allowed_client_redirect_uris.split(",")
        if args.allowed_client_redirect_uris
        else None
    )

    if args.auth_type == "none":
        auth = None
    elif args.auth_type == "static":
        auth = StaticTokenVerifier(
            tokens={
                "test-token": {"client_id": "test-user", "scopes": ["read", "write"]},
                "admin-token": {"client_id": "admin", "scopes": ["admin"]},
            }
        )
    elif args.auth_type == "jwt":
        # Fallback to env vars if not provided via CLI
        jwks_uri = args.token_jwks_uri or os.getenv("FASTMCP_SERVER_AUTH_JWT_JWKS_URI")
        issuer = args.token_issuer or os.getenv("FASTMCP_SERVER_AUTH_JWT_ISSUER")
        audience = args.token_audience or os.getenv("FASTMCP_SERVER_AUTH_JWT_AUDIENCE")
        algorithm = args.token_algorithm
        secret_or_key = args.token_secret or args.token_public_key
        public_key_pem = None

        if not (jwks_uri or secret_or_key):
            logger.error(
                "JWT auth requires either --token-jwks-uri or --token-secret/--token-public-key"
            )
            sys.exit(1)
        if not (issuer and audience):
            logger.error("JWT requires --token-issuer and --token-audience")
            sys.exit(1)

        # Load static public key from file if path is given
        if args.token_public_key and os.path.isfile(args.token_public_key):
            try:
                with open(args.token_public_key, "r") as f:
                    public_key_pem = f.read()
                logger.info(f"Loaded static public key from {args.token_public_key}")
            except Exception as e:
                print(f"Failed to read public key file: {e}")
                logger.error(f"Failed to read public key file: {e}")
                sys.exit(1)
        elif args.token_public_key:
            public_key_pem = args.token_public_key  # Inline PEM

        # Validation: Conflicting options
        if jwks_uri and (algorithm or secret_or_key):
            logger.warning(
                "JWKS mode ignores --token-algorithm and --token-secret/--token-public-key"
            )

        # HMAC mode
        if algorithm and algorithm.startswith("HS"):
            if not secret_or_key:
                logger.error(f"HMAC algorithm {algorithm} requires --token-secret")
                sys.exit(1)
            if jwks_uri:
                logger.error("Cannot use --token-jwks-uri with HMAC")
                sys.exit(1)
            public_key = secret_or_key
        else:
            public_key = public_key_pem

        # Required scopes
        required_scopes = None
        if args.required_scopes:
            required_scopes = [
                s.strip() for s in args.required_scopes.split(",") if s.strip()
            ]

        try:
            auth = JWTVerifier(
                jwks_uri=jwks_uri,
                public_key=public_key,
                issuer=issuer,
                audience=audience,
                algorithm=(
                    algorithm if algorithm and algorithm.startswith("HS") else None
                ),
                required_scopes=required_scopes,
            )
            logger.info(
                "JWTVerifier configured",
                extra={
                    "mode": (
                        "JWKS"
                        if jwks_uri
                        else (
                            "HMAC"
                            if algorithm and algorithm.startswith("HS")
                            else "Static Key"
                        )
                    ),
                    "algorithm": algorithm,
                    "required_scopes": required_scopes,
                },
            )
        except Exception as e:
            print(f"Failed to initialize JWTVerifier: {e}")
            logger.error(f"Failed to initialize JWTVerifier: {e}")
            sys.exit(1)
    elif args.auth_type == "oauth-proxy":
        if not (
            args.oauth_upstream_auth_endpoint
            and args.oauth_upstream_token_endpoint
            and args.oauth_upstream_client_id
            and args.oauth_upstream_client_secret
            and args.oauth_base_url
            and args.token_jwks_uri
            and args.token_issuer
            and args.token_audience
        ):
            print(
                "oauth-proxy requires oauth-upstream-auth-endpoint, oauth-upstream-token-endpoint, "
                "oauth-upstream-client-id, oauth-upstream-client-secret, oauth-base-url, token-jwks-uri, "
                "token-issuer, token-audience"
            )
            logger.error(
                "oauth-proxy requires oauth-upstream-auth-endpoint, oauth-upstream-token-endpoint, "
                "oauth-upstream-client-id, oauth-upstream-client-secret, oauth-base-url, token-jwks-uri, "
                "token-issuer, token-audience",
                extra={
                    "auth_endpoint": args.oauth_upstream_auth_endpoint,
                    "token_endpoint": args.oauth_upstream_token_endpoint,
                    "client_id": args.oauth_upstream_client_id,
                    "base_url": args.oauth_base_url,
                    "jwks_uri": args.token_jwks_uri,
                    "issuer": args.token_issuer,
                    "audience": args.token_audience,
                },
            )
            sys.exit(1)
        token_verifier = JWTVerifier(
            jwks_uri=args.token_jwks_uri,
            issuer=args.token_issuer,
            audience=args.token_audience,
        )
        auth = OAuthProxy(
            upstream_authorization_endpoint=args.oauth_upstream_auth_endpoint,
            upstream_token_endpoint=args.oauth_upstream_token_endpoint,
            upstream_client_id=args.oauth_upstream_client_id,
            upstream_client_secret=args.oauth_upstream_client_secret,
            token_verifier=token_verifier,
            base_url=args.oauth_base_url,
            allowed_client_redirect_uris=allowed_uris,
        )
    elif args.auth_type == "oidc-proxy":
        if not (
            args.oidc_config_url
            and args.oidc_client_id
            and args.oidc_client_secret
            and args.oidc_base_url
        ):
            logger.error(
                "oidc-proxy requires oidc-config-url, oidc-client-id, oidc-client-secret, oidc-base-url",
                extra={
                    "config_url": args.oidc_config_url,
                    "client_id": args.oidc_client_id,
                    "base_url": args.oidc_base_url,
                },
            )
            sys.exit(1)
        auth = OIDCProxy(
            config_url=args.oidc_config_url,
            client_id=args.oidc_client_id,
            client_secret=args.oidc_client_secret,
            base_url=args.oidc_base_url,
            allowed_client_redirect_uris=allowed_uris,
        )
    elif args.auth_type == "remote-oauth":
        if not (
            args.remote_auth_servers
            and args.remote_base_url
            and args.token_jwks_uri
            and args.token_issuer
            and args.token_audience
        ):
            logger.error(
                "remote-oauth requires remote-auth-servers, remote-base-url, token-jwks-uri, token-issuer, token-audience",
                extra={
                    "auth_servers": args.remote_auth_servers,
                    "base_url": args.remote_base_url,
                    "jwks_uri": args.token_jwks_uri,
                    "issuer": args.token_issuer,
                    "audience": args.token_audience,
                },
            )
            sys.exit(1)
        auth_servers = [url.strip() for url in args.remote_auth_servers.split(",")]
        token_verifier = JWTVerifier(
            jwks_uri=args.token_jwks_uri,
            issuer=args.token_issuer,
            audience=args.token_audience,
        )
        auth = RemoteAuthProvider(
            token_verifier=token_verifier,
            authorization_servers=auth_servers,
            base_url=args.remote_base_url,
        )

    # === 2. Build Middleware List ===
    middlewares: List[
        Union[
            UserTokenMiddleware,
            ErrorHandlingMiddleware,
            RateLimitingMiddleware,
            TimingMiddleware,
            LoggingMiddleware,
            JWTClaimsLoggingMiddleware,
            EunomiaMcpMiddleware,
        ]
    ] = [
        ErrorHandlingMiddleware(include_traceback=True, transform_errors=True),
        RateLimitingMiddleware(max_requests_per_second=10.0, burst_capacity=20),
        TimingMiddleware(),
        LoggingMiddleware(),
        JWTClaimsLoggingMiddleware(),
    ]
    if config["enable_delegation"] or args.auth_type == "jwt":
        middlewares.insert(0, UserTokenMiddleware(config=config))  # Must be first

    if args.eunomia_type in ["embedded", "remote"]:
        try:
            from eunomia_mcp import create_eunomia_middleware

            policy_file = args.eunomia_policy_file or "mcp_policies.json"
            eunomia_endpoint = (
                args.eunomia_remote_url if args.eunomia_type == "remote" else None
            )
            eunomia_mw = create_eunomia_middleware(
                policy_file=policy_file, eunomia_endpoint=eunomia_endpoint
            )
            middlewares.append(eunomia_mw)
            logger.info(f"Eunomia middleware enabled ({args.eunomia_type})")
        except Exception as e:
            print(f"Failed to load Eunomia middleware: {e}")
            logger.error("Failed to load Eunomia middleware", extra={"error": str(e)})
            sys.exit(1)

    mcp = FastMCP("Microsoft", auth=auth)
    register_tools(mcp)
    register_prompts(mcp)

    for mw in middlewares:
        mcp.add_middleware(mw)

    print("\nStarting Microsoft MCP Server")
    print(f"  Transport: {args.transport.upper()}")
    print(f"  Auth: {args.auth_type}")
    print(f"  Delegation: {'ON' if config['enable_delegation'] else 'OFF'}")
    print(f"  Eunomia: {args.eunomia_type}")

    if args.transport == "stdio":
        mcp.run(transport="stdio")
    elif args.transport == "streamable-http":
        mcp.run(transport="streamable-http", host=args.host, port=args.port)
    elif args.transport == "sse":
        mcp.run(transport="sse", host=args.host, port=args.port)
    else:
        logger.error("Invalid transport", extra={"transport": args.transport})
        sys.exit(1)


def usage():
    print(
        f"Microsoft Agent ({__version__}): Microsoft MCP Server\n\n"
        "Usage:\n"
        "-t | --transport                   [ Transport method: 'stdio', 'streamable-http', or 'sse' [legacy] (default: stdio) ]\n"
        "-s | --host                        [ Host address for HTTP transport (default: 0.0.0.0) ]\n"
        "-p | --port                        [ Port number for HTTP transport (default: 8000) ]\n"
        "--auth-type                        [ Authentication type for MCP server: 'none' (disabled), 'static' (internal), 'jwt' (external token verification), 'oauth-proxy', 'oidc-proxy', 'remote-oauth' (external) (default: none) ]\n"
        "--token-jwks-uri                   [ JWKS URI for JWT verification ]\n"
        "--token-issuer                     [ Issuer for JWT verification ]\n"
        "--token-audience                   [ Audience for JWT verification ]\n"
        "--token-algorithm                  [ JWT signing algorithm (required for HMAC or static key). Auto-detected for JWKS. ]\n"
        "--token-secret                     [ Shared secret for HMAC (HS*) or PEM public key for static asymmetric verification. ]\n"
        "--token-public-key                 [ Path to PEM public key file or inline PEM string (for static asymmetric keys). ]\n"
        "--required-scopes                  [ Comma-separated list of required scopes (e.g., microsoft.read,microsoft.write). ]\n"
        "--oauth-upstream-auth-endpoint     [ Upstream authorization endpoint for OAuth Proxy ]\n"
        "--oauth-upstream-token-endpoint    [ Upstream token endpoint for OAuth Proxy ]\n"
        "--oauth-upstream-client-id         [ Upstream client ID for OAuth Proxy ]\n"
        "--oauth-upstream-client-secret     [ Upstream client secret for OAuth Proxy ]\n"
        "--oauth-base-url                   [ Base URL for OAuth Proxy ]\n"
        "--oidc-config-url                  [ OIDC configuration URL ]\n"
        "--oidc-client-id                   [ OIDC client ID ]\n"
        "--oidc-client-secret               [ OIDC client secret ]\n"
        "--oidc-base-url                    [ Base URL for OIDC Proxy ]\n"
        "--remote-auth-servers              [ Comma-separated list of authorization servers for Remote OAuth ]\n"
        "--remote-base-url                  [ Base URL for Remote OAuth ]\n"
        "--allowed-client-redirect-uris     [ Comma-separated list of allowed client redirect URIs ]\n"
        "--eunomia-type                     [ Eunomia authorization type: 'none' (disabled), 'embedded' (built-in), 'remote' (external) (default: none) ]\n"
        "--eunomia-policy-file              [ Policy file for embedded Eunomia (default: mcp_policies.json) ]\n"
        "--eunomia-remote-url               [ URL for remote Eunomia server ]\n"
        "--enable-delegation                [ Enable OIDC token delegation ]\n"
        "--audience                         [ Audience for the delegated token ]\n"
        "--delegated-scopes                 [ Scopes for the delegated token (space-separated) ]\n"
        "--openapi-file                     [ Path to the OpenAPI JSON file to import additional tools from ]\n"
        "--openapi-base-url                 [ Base URL for the OpenAPI client (overrides instance URL) ]\n"
        "--openapi-use-token                [ Use the incoming Bearer token (from MCP request) to authenticate OpenAPI import ]\n"
        "--openapi-username                 [ Username for basic auth during OpenAPI import ]\n"
        "--openapi-password                 [ Password for basic auth during OpenAPI import ]\n"
        "--openapi-client-id                [ OAuth client ID for OpenAPI import ]\n"
        "--openapi-client-secret            [ OAuth client secret for OpenAPI import ]\n"
        "\n"
        "Examples:\n"
        "  [Simple]  microsoft-mcp \n"
        '  [Complex] microsoft-mcp --transport "value" --host "value" --port "value" --auth-type "value" --token-jwks-uri "value" --token-issuer "value" --token-audience "value" --token-algorithm "value" --token-secret "value" --token-public-key "value" --required-scopes "value" --oauth-upstream-auth-endpoint "value" --oauth-upstream-token-endpoint "value" --oauth-upstream-client-id "value" --oauth-upstream-client-secret "value" --oauth-base-url "value" --oidc-config-url "value" --oidc-client-id "value" --oidc-client-secret "value" --oidc-base-url "value" --remote-auth-servers "value" --remote-base-url "value" --allowed-client-redirect-uris "value" --eunomia-type "value" --eunomia-policy-file "value" --eunomia-remote-url "value" --enable-delegation --audience "value" --delegated-scopes "value" --openapi-file "value" --openapi-base-url "value" --openapi-use-token --openapi-username "value" --openapi-password "value" --openapi-client-id "value" --openapi-client-secret "value"\n'
    )


if __name__ == "__main__":
    microsoft_mcp()
