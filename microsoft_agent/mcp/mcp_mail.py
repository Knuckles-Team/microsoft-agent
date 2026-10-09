"""MCP tools for mail operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_connector_sdk.mcp.action_dispatch import resolve_action
from agent_connector_sdk.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

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
    "list_folder_files",
    "list_chat_messages",
    "get_chat_message",
    "send_chat_message",
    "list_channel_messages",
    "get_channel_message",
    "send_channel_message",
    "list_channel_message_replies",
    "reply_to_channel_message",
    "list_chat_message_replies",
    "reply_to_chat_message",
)


def register_mail_tools(mcp: FastMCP):
    @mcp.tool(tags={"mail"})
    async def microsoft_mail(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_mail_messages', 'list_mail_folders', 'list_mail_folder_messages', 'get_mail_message', 'send_mail', 'list_shared_mailbox_messages', 'list_shared_mailbox_folder_messages', 'get_shared_mailbox_message', 'send_shared_mailbox_mail', 'create_draft_email', 'delete_mail_message', 'move_mail_message', 'update_mail_message', 'add_mail_attachment', 'list_mail_attachments', 'get_mail_attachment', 'delete_mail_attachment', 'list_folder_files', 'list_chat_messages', 'get_chat_message', 'send_chat_message', 'list_channel_messages', 'get_channel_message', 'send_channel_message', 'list_channel_message_replies', 'reply_to_channel_message', 'list_chat_message_replies', 'reply_to_chat_message'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft mail operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _MAIL_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
