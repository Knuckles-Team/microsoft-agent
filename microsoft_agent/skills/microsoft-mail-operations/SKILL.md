---
name: microsoft-mail-operations
skill_type: skill
description: >-
  Outlook / Exchange mail operations on Microsoft Graph via the microsoft-agent
  MCP server — list and read messages, send mail and drafts, manage folders, and
  handle attachments with the domain-typed `microsoft_mail` tool. Use when the
  agent must triage a mailbox, read a message by id, send or draft an email, or
  pull attachment bytes into the knowledge graph. Do NOT use for calendar events
  (use microsoft-calendar-scheduling), Teams chat/channel posts
  (use microsoft-teams-messaging), or OneDrive/SharePoint files.
license: MIT
tags: [microsoft, outlook, mail, graph, mcp]
metadata:
  author: Genius
  version: '0.1.0'
---
# Microsoft Mail Operations

Domain-typed access to Outlook mail through Microsoft Graph (`/me/messages`,
`/me/mailFolders`). Prefer the `microsoft_mail` tool over raw Graph calls — it
carries the message/folder/attachment action set and returns Graph-shaped records.

## When to use
- Triage a mailbox: list messages, filter by folder, search subject/body.
- Read a single message by `id` (headers, body, attachments).
- Send a mail, create a draft, move/update/delete a message.
- List or fetch attachments (and store their bytes in the KG).

## When NOT to use
- Calendar events / meetings → `microsoft-calendar-scheduling`.
- Teams 1:1 chat or channel posts → `microsoft-teams-messaging`.
- OneDrive / SharePoint drive items → the `microsoft_files` tool.
- Directory users/groups → the `microsoft_user` / `microsoft_groups` tools.

## Prerequisites & environment
Connect via the `mcp-client` skill against the **`microsoft-agent`** MCP server.
Auth is interactive MSAL (device/auth-code) — run the `microsoft_auth` tool
(`login`) first. Requires the `Mail.ReadWrite` (and `Mail.Send`) Graph scopes.

| Variable | Required | Notes |
|----------|----------|-------|
| `OIDC_CLIENT_ID` | optional | Azure AD app (client) id; a default is baked in |
| `TESTING` | optional | Set truthy to skip the global auth manager |

`MCP_TOOL_MODE` (`condensed`|`verbose`|`both`) selects the condensed surface
(used below) vs. the 1:1 verbose tools.

## Tools & actions
Prefer the **condensed** `microsoft_mail` tool; it takes `action` + a `params_json`
**JSON string** whose keys are passed straight to the client method (usually inside
a `params` object of OData query options).

| Condensed tool | Key actions |
|----------------|-------------|
| `microsoft_mail` | `list_mail_messages`, `list_mail_folders`, `list_mail_folder_messages`, `get_mail_message`, `send_mail`, `create_draft_email`, `move_mail_message`, `update_mail_message`, `delete_mail_message`, `list_mail_attachments`, `get_mail_attachment` |

## Recipes (`params_json`)
List the 25 most recent messages (few fields):
```json
{"params":{"$top":25,"$select":"id,subject,from,receivedDateTime,hasAttachments","$orderby":"receivedDateTime desc"}}
```
Search the mailbox by subject:
```json
{"params":{"$search":"\"quarterly report\"","$top":10}}
```
Get one message by id:
```json
{"message_id":"AAMkAG...","params":{"$select":"subject,body,from,toRecipients"}}
```
List a message's attachments:
```json
{"message_id":"AAMkAG..."}
```

## Gotchas
- `params_json` is a **string** of JSON, not an object — serialize it.
- OData query options go **inside** a `params` object (`$select`, `$filter`,
  `$search`, `$top`, `$orderby`); `$search` requires quoting and adds a
  `ConsistencyLevel: eventual` requirement server-side.
- Responses are raw Graph JSON: the collection is under the `value` array, not
  `data`.
- `$select` a small field set — full message bodies are large and slow.
- Attachment bytes arrive base64 in `contentBytes` on `fileAttachment` records.

## Related
- **Native KG ingestion:** the `microsoft_ingest_records` tool (`kind:"messages"`)
  maps messages → `:Message` (+ `:Person` sender/recipients + `:Document` body)
  nodes; `microsoft_agent.kg_media.ingest_attachment` stores attachment bytes as a
  `:Blob` / `:MediaAsset`. Ingestion only — not part of the triage surface.
- **Source preset:** `microsoft-mail` in `connectors/mcp_source_presets.json` syncs
  mail into the KG as `email` documents.
