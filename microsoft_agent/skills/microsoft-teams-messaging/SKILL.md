---
name: microsoft-teams-messaging
skill_type: skill
description: >-
  Microsoft Teams chat and channel messaging on Microsoft Graph via the
  microsoft-agent MCP server — list and read 1:1/group chat messages and channel
  posts, send messages, and reply in threads using the `microsoft_chat` and
  `microsoft_teams` tools. Use when the agent must read a Teams conversation,
  post a message to a chat or channel, or reply to a thread. Do NOT use for
  Outlook mail (use microsoft-mail-operations) or calendar events
  (use microsoft-calendar-scheduling).
license: MIT
tags: [microsoft, teams, chat, channel, mcp]
metadata:
  author: Genius
  version: '0.1.0'
---
# Microsoft Teams Messaging

Domain-typed access to Teams conversations through Microsoft Graph (`/chats`,
`/teams/{id}/channels/{id}/messages`). Chat/channel message verbs live on the
`microsoft_mail` action router (Graph co-locates them), while team and chat
container reads use the `microsoft_teams` / `microsoft_chat` tools.

## When to use
- Read messages in a 1:1 or group chat, or a channel's posts.
- Send a chat message or a channel message.
- List replies to a message, or reply to a thread.
- Resolve a team, its channels, or a chat container.

## When NOT to use
- Outlook mailbox messages → `microsoft-mail-operations`.
- Calendar events / meetings → `microsoft-calendar-scheduling`.
- Directory group membership management → the `microsoft_groups` tool.

## Prerequisites & environment
Connect via the `mcp-client` skill against the **`microsoft-agent`** MCP server.
Run `microsoft_auth` (`login`) first. Requires `Chat.Read`/`ChatMessage.Read.All`
and `ChannelMessage.Read.All` (plus send scopes) Graph permissions.

| Variable | Required | Notes |
|----------|----------|-------|
| `OIDC_CLIENT_ID` | optional | Azure AD app (client) id; a default is baked in |
| `TESTING` | optional | Set truthy to skip the global auth manager |

## Tools & actions
`action` + a `params_json` **JSON string**.

| Condensed tool | Key actions |
|----------------|-------------|
| `microsoft_mail` | `list_chat_messages`, `get_chat_message`, `send_chat_message`, `list_channel_messages`, `get_channel_message`, `send_channel_message`, `list_chat_message_replies`, `reply_to_chat_message` |
| `microsoft_teams` | `get_team`, `get_team_channel` |
| `microsoft_chat` | `get_chat` |

## Recipes (`params_json`)
List messages in a chat:
```json
{"chat_id":"19:abc...@thread.v2","params":{"$top":20,"$orderby":"createdDateTime desc"}}
```
List a channel's messages:
```json
{"team_id":"<team-guid>","channel_id":"19:def...@thread.tacv2","params":{"$top":20}}
```
Send a channel message:
```json
{"team_id":"<team-guid>","channel_id":"19:def...@thread.tacv2","body":{"body":{"content":"Deploy is green ✅"}}}
```

## Gotchas
- `params_json` is a **string** of JSON — serialize it.
- Chat/channel message verbs live under the **`microsoft_mail`** action set (Graph
  models them as messages), not under `microsoft_teams`.
- `chat_id` / `team_id` / `channel_id` are Graph thread ids (`19:...@thread.*`),
  not display names — resolve them first via `list_chats` / `list_joined_teams`.
- Responses are raw Graph JSON — the collection is under `value`.

## Related
- **Native KG ingestion:** the `microsoft_ingest_records` tool (`kind:"messages"`)
  maps chat/channel messages → `:Message` (+ `:Person` sender) nodes, the same
  mapper used for mail.
- **Composed by:** the `microsoft-mail-operations` skill shares the underlying
  `microsoft_mail` tool for the message verbs.
