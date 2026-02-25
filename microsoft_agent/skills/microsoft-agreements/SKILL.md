---
name: microsoft-agreements
description: "Microsoft 365 Agreements — Terms-of-Use Agreements Management"
tags: [agreements]
---

# Microsoft 365 Agreements

Manage terms-of-use agreements for the tenant.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_agreement` | Create a terms-of-use agreement |
| `delete_agreement` | Delete an agreement |
| `get_agreement` | Get a specific agreement |
| `list_agreements` | List terms-of-use agreements |

## Required Permissions
- `Agreement.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
