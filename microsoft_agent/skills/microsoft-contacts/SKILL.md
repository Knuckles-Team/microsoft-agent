---
name: microsoft-contacts
skill_type: skill
description: >-
  Outlook personal contacts on Microsoft Graph: list, read, create, update and
  delete contacts. Use when the agent must look up or maintain a contact. Do NOT
  use for directory users and groups (use microsoft-identity-directory).
license: MIT
tags: [microsoft, contacts, outlook]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Outlook contacts

Personal contacts only; tenant users live in the directory.

## When to use
- Look up a contact.
- Create, update or delete a contact.

## When NOT to use
- Tenant users and groups: use microsoft-identity-directory.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (5)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_outlook_contact` | W | data, params |
| `microsoft_delete_outlook_contact` | D | confirm, contact_id, params |
| `microsoft_get_outlook_contact` | R | contact_id, params |
| `microsoft_list_outlook_contacts` | R | params |
| `microsoft_update_outlook_contact` | W | contact_id, data, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `contacts-outlook-contact`: `microsoft_get_outlook_contact` -> `microsoft_list_outlook_contacts` -> `microsoft_create_outlook_contact` -> `microsoft_update_outlook_contact` -> `microsoft_delete_outlook_contact`
