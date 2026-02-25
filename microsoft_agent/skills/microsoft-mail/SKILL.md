---
name: microsoft-mail
description: "Microsoft 365 Mail — Email Messages, Folders, Attachments, Drafts & Shared Mailboxes"
tags: [mail]
---

# Microsoft 365 Mail

Manage email messages, folders, attachments, shared mailboxes, drafts, and send mail.

## Available Tools

| Tool | Description |
|------|-------------|
| `add_mail_attachment` | add_mail_attachment: POST /me/messages/{message-id}/attachments |
| `create_draft_email` | create_draft_email: POST /me/messages |
| `delete_mail_attachment` | delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id} |
| `delete_mail_message` | delete_mail_message: DELETE /me/messages/{message-id} |
| `get_channel_message` | get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id} |
| `get_chat_message` | get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id} |
| `get_mail_attachment` | get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id} |
| `get_mail_message` | get_mail_message: GET /me/messages/{message-id} |
| `get_root_folder` | get_root_folder: GET /drives/{drive-id}/root |
| `get_shared_mailbox_message` | get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id} |
| `list_channel_messages` | list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages |
| `list_chat_message_replies` | list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `list_chat_messages` | list_chat_messages: GET /chats/{chat-id}/messages |
| `list_folder_files` | list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children |
| `list_mail_attachments` | list_mail_attachments: GET /me/messages/{message-id}/attachments |
| `list_mail_folder_messages` | list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages |
| `list_mail_messages` | list_mail_messages: GET /me/messages |
| `list_mail_folders` | list_mail_folders: GET /me/mailFolders |
| `list_shared_mailbox_folder_messages` | list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages |
| `list_shared_mailbox_messages` | list_shared_mailbox_messages: GET /users/{user-id}/messages |
| `move_mail_message` | move_mail_message: POST /me/messages/{message-id}/move |
| `reply_to_chat_message` | reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `send_channel_message` | send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages |
| `send_chat_message` | send_chat_message: POST /chats/{chat-id}/messages |
| `send_mail` | TIP: CRITICAL: Do not try to guess the email address of the recipients |
| `send_shared_mailbox_mail` | TIP: CRITICAL: Do not try to guess the email address of the recipients |
| `update_mail_message` | update_mail_message: PATCH /me/messages/{message-id} |

## Required Permissions
- `Mail.ReadWrite, Mail.Send`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
