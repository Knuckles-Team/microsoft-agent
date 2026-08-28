"""MCP tools for mail operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp.action_dispatch import resolve_action
from agent_utilities.mcp.concurrency import invoke_client_method
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

        if action == "list_mail_messages":
            return await invoke_client_method(client.list_mail_messages, **kwargs)
        if action == "list_mail_folders":
            return await invoke_client_method(client.list_mail_folders, **kwargs)
        if action == "list_mail_folder_messages":
            return await invoke_client_method(
                client.list_mail_folder_messages, **kwargs
            )
        if action == "get_mail_message":
            return await invoke_client_method(client.get_mail_message, **kwargs)
        if action == "send_mail":
            return await invoke_client_method(client.send_mail, **kwargs)
        if action == "list_shared_mailbox_messages":
            return await invoke_client_method(
                client.list_shared_mailbox_messages, **kwargs
            )
        if action == "list_shared_mailbox_folder_messages":
            return await invoke_client_method(
                client.list_shared_mailbox_folder_messages, **kwargs
            )
        if action == "get_shared_mailbox_message":
            return await invoke_client_method(
                client.get_shared_mailbox_message, **kwargs
            )
        if action == "send_shared_mailbox_mail":
            return await invoke_client_method(client.send_shared_mailbox_mail, **kwargs)
        if action == "create_draft_email":
            return await invoke_client_method(client.create_draft_email, **kwargs)
        if action == "delete_mail_message":
            return await invoke_client_method(client.delete_mail_message, **kwargs)
        if action == "move_mail_message":
            return await invoke_client_method(client.move_mail_message, **kwargs)
        if action == "update_mail_message":
            return await invoke_client_method(client.update_mail_message, **kwargs)
        if action == "add_mail_attachment":
            return await invoke_client_method(client.add_mail_attachment, **kwargs)
        if action == "list_mail_attachments":
            return await invoke_client_method(client.list_mail_attachments, **kwargs)
        if action == "get_mail_attachment":
            return await invoke_client_method(client.get_mail_attachment, **kwargs)
        if action == "delete_mail_attachment":
            return await invoke_client_method(client.delete_mail_attachment, **kwargs)
        if action == "list_folder_files":
            return await invoke_client_method(client.list_folder_files, **kwargs)
        if action == "list_chat_messages":
            return await invoke_client_method(client.list_chat_messages, **kwargs)
        if action == "get_chat_message":
            return await invoke_client_method(client.get_chat_message, **kwargs)
        if action == "send_chat_message":
            return await invoke_client_method(client.send_chat_message, **kwargs)
        if action == "list_channel_messages":
            return await invoke_client_method(client.list_channel_messages, **kwargs)
        if action == "get_channel_message":
            return await invoke_client_method(client.get_channel_message, **kwargs)
        if action == "send_channel_message":
            return await invoke_client_method(client.send_channel_message, **kwargs)
        if action == "list_channel_message_replies":
            return await invoke_client_method(
                client.list_channel_message_replies, **kwargs
            )
        if action == "reply_to_channel_message":
            return await invoke_client_method(client.reply_to_channel_message, **kwargs)
        if action == "list_chat_message_replies":
            return await invoke_client_method(
                client.list_chat_message_replies, **kwargs
            )
        if action == "reply_to_chat_message":
            return await invoke_client_method(client.reply_to_chat_message, **kwargs)
        raise ValueError(f"Unknown action: {action}")
