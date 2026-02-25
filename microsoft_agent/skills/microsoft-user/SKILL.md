---
name: microsoft-user
description: "Microsoft 365 User — User Profile, Mail Operations, Meetings & Group Membership"
tags: [user]
---

# Microsoft 365 User

Manage user profiles, mail operations, meetings, and group membership.

## Available Tools

| Tool | Description |
|------|-------------|
| `add_group_member` | Add a member to a group |
| `add_mail_attachment` | add_mail_attachment: POST /me/messages/{message-id}/attachments |
| `delete_mail_attachment` | delete_mail_attachment: DELETE /me/messages/{message-id}/attachments/{attachment-id} |
| `delete_mail_message` | delete_mail_message: DELETE /me/messages/{message-id} |
| `find_meeting_times` | find_meeting_times: POST /me/findMeetingTimes |
| `get_channel_message` | get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id} |
| `get_chat_message` | get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id} |
| `get_current_user` | get_current_user: GET /me |
| `get_mail_attachment` | get_mail_attachment: GET /me/messages/{message-id}/attachments/{attachment-id} |
| `get_mail_message` | get_mail_message: GET /me/messages/{message-id} |
| `get_me` | get_me: GET /me |
| `get_shared_mailbox_message` | get_shared_mailbox_message: GET /users/{user-id}/messages/{message-id} |
| `list_channel_messages` | list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages |
| `list_chat_message_replies` | list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `list_chat_messages` | list_chat_messages: GET /chats/{chat-id}/messages |
| `list_group_members` | Get a list of the group |
| `list_group_owners` | Get owners of a group |
| `list_mail_attachments` | list_mail_attachments: GET /me/messages/{message-id}/attachments |
| `list_mail_folder_messages` | list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages |
| `list_mail_messages` | list_mail_messages: GET /me/messages |
| `list_shared_mailbox_folder_messages` | list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages |
| `list_shared_mailbox_messages` | list_shared_mailbox_messages: GET /users/{user-id}/messages |
| `list_team_members` | list_team_members: GET /teams/{team-id}/members |
| `list_users` | list_users: GET /users |
| `move_mail_message` | move_mail_message: POST /me/messages/{message-id}/move |
| `remove_group_member` | Remove a member from a group |
| `reply_to_chat_message` | reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `send_channel_message` | send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages |
| `send_chat_message` | send_chat_message: POST /chats/{chat-id}/messages |
| `update_mail_message` | update_mail_message: PATCH /me/messages/{message-id} |

## Required Permissions
- `User.Read, Mail.ReadWrite, Chat.Read, Group.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
