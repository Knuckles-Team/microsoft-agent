---
name: microsoft-auth
description: "Microsoft 365 Auth — Authentication & Session Management"
---

# Microsoft 365 Auth

Manage authentication operations including login, logout, session verification, and account listing.

## Available Tools

| Tool | Description |
|------|-------------|
| `list_accounts` | List all available Microsoft accounts |
| `login` | Authenticate with Microsoft using device code flow |
| `logout` | Log out from Microsoft account |
| `verify_login` | Check current Microsoft authentication status |

## Required Permissions
- `User.Read`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
