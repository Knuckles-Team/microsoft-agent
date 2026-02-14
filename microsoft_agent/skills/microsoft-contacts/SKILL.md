---
name: microsoft-contacts
description: "Microsoft 365 Contacts — Outlook Contact Management"
---

# Microsoft 365 Contacts

Manage Outlook contacts — create, read, update, and delete.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_outlook_contact` | create_outlook_contact: POST /me/contacts |
| `delete_outlook_contact` | delete_outlook_contact: DELETE /me/contacts/{contact-id} |
| `get_outlook_contact` | get_outlook_contact: GET /me/contacts/{contact-id} |
| `list_outlook_contacts` | list_outlook_contacts: GET /me/contacts |
| `update_outlook_contact` | update_outlook_contact: PATCH /me/contacts/{contact-id} |

## Required Permissions
- `Contacts.ReadWrite`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
