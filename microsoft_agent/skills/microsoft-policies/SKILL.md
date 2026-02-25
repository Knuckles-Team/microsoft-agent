---
name: microsoft-policies
description: "Microsoft 365 Policies — Authorization, Token & Permission Grant Policies"
tags: [policies]
---

# Microsoft 365 Policies

Manage authorization policies, token policies, permission grant policies, and admin consent policies.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_admin_consent_policy` | Get the admin consent request policy |
| `get_authorization_policy` | Get the tenant authorization policy |
| `list_permission_grant_policies` | List permission grant policies |
| `list_token_issuance_policies` | List token issuance policies |
| `list_token_lifetime_policies` | List token lifetime policies |

## Required Permissions
- `Policy.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
