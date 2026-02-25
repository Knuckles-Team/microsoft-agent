---
name: microsoft-chat
description: "Microsoft 365 Chat — Chats, Messages, Replies & Group Conversations"
tags: [chat]
---

# Microsoft 365 Chat

Manage chats, messages, replies, and group conversations.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_chat` | get_chat: GET /chats/{chat-id} |
| `get_chat_message` | get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id} |
| `list_chat_message_replies` | list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `list_chat_messages` | list_chat_messages: GET /chats/{chat-id}/messages |
| `list_chats` | list_chats: GET /me/chats |
| `list_group_conversations` | List conversations in a Microsoft 365 group |
| `reply_to_chat_message` | reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `send_chat_message` | send_chat_message: POST /chats/{chat-id}/messages |

## Required Permissions
- `Chat.Read, ChatMessage.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
