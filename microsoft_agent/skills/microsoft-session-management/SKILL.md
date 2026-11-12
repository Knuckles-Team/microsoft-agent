---
name: microsoft-session-management
skill_type: skill
description: >-
  Sign in to Microsoft Graph, verify the session, list cached accounts, sign out
  and discover tools. Use at the start of any Microsoft task, or when a call
  reports an expired session. Do NOT use for the work itself (use the domain
  skills).
license: MIT
tags: [microsoft, auth, session]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft session management

Interactive and service sign-in for the Microsoft Graph connector. Always verify the session before a write.

## When to use
- Start a session or confirm which account is active.
- Sign out after a privileged task.
- Find the right tool by keyword.

## When NOT to use
- Reading mail, calendar, files or directory data: use the domain skills.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (7)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `login_power_platform` | W | none |
| `login_windows_companion` | W | none |
| `microsoft_list_accounts` | R | none |
| `microsoft_login` | W | force |
| `microsoft_logout` | W | none |
| `microsoft_search_tools` | R | limit, query |
| `microsoft_verify_login` | W | none |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `session-login`: `microsoft_login` -> `microsoft_verify_login`
- `session-read-records`: `microsoft_list_accounts` -> `microsoft_search_tools`
- `session-change-records`: `login_power_platform` -> `login_windows_companion` -> `microsoft_logout`
