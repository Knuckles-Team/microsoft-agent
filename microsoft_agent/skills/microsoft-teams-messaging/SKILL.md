---
name: microsoft-teams-messaging
skill_type: skill
description: >-
  Microsoft Teams teams, channels, chats, messages, replies, presence and call
  records on Microsoft Graph. Use when the agent must read or post in a chat or
  channel, reply to a message, list joined teams, or check presence. Do NOT use
  for mail (use microsoft-mail-operations) or files (use microsoft-files-
  sharepoint).
license: MIT
tags: [microsoft, teams, chat, graph, mcp]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Teams collaboration

Typed Teams operations. Posting a message is a visible side effect: confirm the target chat or channel first.

## When to use
- List joined teams, channels, chats and members.
- Read channel and chat messages and replies.
- Send a message or reply.
- Check presence; read call records.

## When NOT to use
- Files shared in Teams: use microsoft-files-sharepoint.
- Mail: use microsoft-mail-operations.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (23)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_get_call_record` | R | call_id, params |
| `microsoft_get_channel_message` | R | channel_id, chatMessage_id, params, team_id |
| `microsoft_get_chat` | R | chat_id, params |
| `microsoft_get_chat_message` | R | chatMessage_id, chat_id, params |
| `microsoft_get_my_presence` | R | params |
| `microsoft_get_presence` | R | params, user_id |
| `microsoft_get_team` | R | params, team_id |
| `microsoft_get_team_channel` | R | channel_id, params, team_id |
| `microsoft_get_teams_user_activity` | R | params, period |
| `microsoft_list_call_records` | R | params |
| `microsoft_list_channel_message_replies` | R | channel_id, chatMessage_id, params, team_id |
| `microsoft_list_channel_messages` | R | channel_id, params, team_id |
| `microsoft_list_chat_message_replies` | R | chatMessage_id, chat_id, params |
| `microsoft_list_chat_messages` | R | chat_id, params |
| `microsoft_list_chats` | R | params |
| `microsoft_list_joined_teams` | R | params |
| `microsoft_list_presences` | R | params |
| `microsoft_list_team_channels` | R | params, team_id |
| `microsoft_list_team_members` | R | params, team_id |
| `microsoft_reply_to_channel_message` | W | channel_id, chatMessage_id, data, params, team_id |
| `microsoft_reply_to_chat_message` | W | chatMessage_id, chat_id, data, params |
| `microsoft_send_channel_message` | W | channel_id, data, params, team_id |
| `microsoft_send_chat_message` | W | chat_id, data, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `teams-call-record`: `microsoft_get_call_record` -> `microsoft_list_call_records`
- `teams-channel-message`: `microsoft_get_channel_message` -> `microsoft_list_channel_message_replies` -> `microsoft_list_channel_messages` -> `microsoft_send_channel_message` -> `microsoft_reply_to_channel_message`
- `teams-chat`: `microsoft_get_chat` -> `microsoft_list_chats`
- `teams-chat-message`: `microsoft_get_chat_message` -> `microsoft_list_chat_message_replies` -> `microsoft_list_chat_messages` -> `microsoft_send_chat_message` -> `microsoft_reply_to_chat_message`
- `teams-presence`: `microsoft_get_my_presence` -> `microsoft_get_presence` -> `microsoft_list_presences`
- `teams-team-channel`: `microsoft_get_team_channel` -> `microsoft_list_team_channels`
- `teams-read-records`: `microsoft_get_team` -> `microsoft_get_teams_user_activity` -> `microsoft_list_joined_teams` -> `microsoft_list_team_members`
