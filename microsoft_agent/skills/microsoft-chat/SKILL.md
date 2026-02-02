---
name: microsoft-chat
description: "Generated skill for chat operations. Contains 7 tools."
---

### Overview
This skill handles operations related to chat.

### Available Tools
- `list_chats`: list_chats: GET /me/chats
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `get_chat`: get_chat: GET /chats/{chat-id}
  - **Parameters**:
    - `chat_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_chat_messages`: list_chat_messages: GET /chats/{chat-id}/messages
  - **Parameters**:
    - `chat_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_chat_message`: get_chat_message: GET /chats/{chat-id}/messages/{chatMessage-id}
  - **Parameters**:
    - `chat_id` (str)
    - `chatMessage_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `send_chat_message`: send_chat_message: POST /chats/{chat-id}/messages
  - **Parameters**:
    - `chat_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `list_chat_message_replies`: list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies
  - **Parameters**:
    - `chat_id` (str)
    - `chatMessage_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `reply_to_chat_message`: reply_to_chat_message: POST /chats/{chat-id}/messages/{chatMessage-id}/replies
  - **Parameters**:
    - `chat_id` (str)
    - `chatMessage_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
