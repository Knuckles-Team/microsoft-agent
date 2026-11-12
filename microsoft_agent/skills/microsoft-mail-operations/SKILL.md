---
name: microsoft-mail-operations
skill_type: skill
description: >-
  Outlook and Exchange mail through Microsoft Graph: list and read messages,
  folders and attachments, send mail and drafts, move or delete messages, read
  shared mailboxes and mailbox reports. Use when the agent must triage a
  mailbox, send or draft an email, or fetch an attachment. Do NOT use for
  calendar events (use microsoft-calendar-scheduling) or Teams messages (use
  microsoft-teams-messaging).
license: MIT
tags: [microsoft, outlook, mail, graph, mcp]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft mail operations

Typed mail operations over `/me/messages`, `/me/mailFolders` and shared mailboxes. Needs the Mail.ReadWrite and Mail.Send scopes.

## When to use
- Triage a mailbox: list folders, list messages, read one message.
- Draft, send, move or delete a message.
- Read or add an attachment.
- Read a shared mailbox or the mailbox usage and email activity reports.

## When NOT to use
- Calendar events: use microsoft-calendar-scheduling.
- Teams chat and channel posts: use microsoft-teams-messaging.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (19)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_add_mail_attachment` | W | data, message_id, params |
| `microsoft_create_draft_email` | W | data, params |
| `microsoft_delete_mail_attachment` | D | attachment_id, confirm, message_id, params |
| `microsoft_delete_mail_message` | D | confirm, message_id, params |
| `microsoft_get_email_activity_report` | R | params, period |
| `microsoft_get_mail_attachment` | R | attachment_id, message_id, params |
| `microsoft_get_mail_message` | R | message_id, params |
| `microsoft_get_mailbox_usage_report` | R | params, period |
| `microsoft_get_shared_mailbox_message` | R | message_id, params, user_id |
| `microsoft_list_mail_attachments` | R | message_id, params |
| `microsoft_list_mail_folder_messages` | R | mailFolder_id, params |
| `microsoft_list_mail_folders` | R | params |
| `microsoft_list_mail_messages` | R | params |
| `microsoft_list_shared_mailbox_folder_messages` | R | mailFolder_id, params, user_id |
| `microsoft_list_shared_mailbox_messages` | R | params, user_id |
| `microsoft_move_mail_message` | W | data, message_id, params |
| `microsoft_send_mail` | W | data, params |
| `microsoft_send_shared_mailbox_mail` | W | data, params, user_id |
| `microsoft_update_mail_message` | W | data, message_id, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `mail-attachment`: `microsoft_get_mail_attachment` -> `microsoft_list_mail_attachments` -> `microsoft_add_mail_attachment` -> `microsoft_delete_mail_attachment`
- `mail-folder`: `microsoft_list_mail_folder_messages` -> `microsoft_list_mail_folders`
- `mail-message`: `microsoft_get_mail_message` -> `microsoft_list_mail_messages` -> `microsoft_move_mail_message` -> `microsoft_update_mail_message` -> `microsoft_delete_mail_message`
- `mail-shared-mailbox`: `microsoft_get_shared_mailbox_message` -> `microsoft_list_shared_mailbox_folder_messages` -> `microsoft_list_shared_mailbox_messages` -> `microsoft_send_shared_mailbox_mail`
- `mail-read-records`: `microsoft_get_email_activity_report` -> `microsoft_get_mailbox_usage_report`
- `mail-change-records`: `microsoft_create_draft_email` -> `microsoft_send_mail`
