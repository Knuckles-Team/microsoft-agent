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
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Steps

1. List cached accounts with `microsoft_list_accounts`.
   Verify whether the intended account is present.
2. Sign in with `microsoft_login` when no account is signed in.
   Verify the sign-in prompt or device code completes.
3. Confirm the session with `microsoft_verify_login`.
   Verify the response reports a signed-in account and tenant.
4. Discover tools for the task with `microsoft_search_tools` (query, limit).
   Verify the tool you need appears by name.

## Failure recovery

| Failure | Cause | Recovery |
|---|---|---|
| Login does not complete | The device code expired or was not entered | Call `microsoft_login` with `force` set to true and finish the new prompt |
| Authentication error or expired session | The Graph session is missing or expired | Call `microsoft_login`, call `microsoft_verify_login`, and resubmit the original call (see microsoft-session-management) |
| Wrong account in use | A different cached account is active | Call `microsoft_logout`, then `microsoft_login`, and confirm with `microsoft_verify_login` |

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
