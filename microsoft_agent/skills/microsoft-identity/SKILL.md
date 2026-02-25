---
name: microsoft-identity
description: "Microsoft 365 Identity — Invitations, Conditional Access, Access Reviews & Lifecycle Workflows"
tags: [identity]
---

# Microsoft 365 Identity

Manage identity operations including invitations, conditional access, access reviews, entitlement access packages, and lifecycle workflows.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_conditional_access_policy` | Create a conditional access policy |
| `create_invitation` | Create an invitation for an external / guest user |
| `delete_conditional_access_policy` | Delete a conditional access policy |
| `get_access_review` | Get a specific access review definition |
| `get_conditional_access_policy` | Get a specific conditional access policy |
| `list_access_reviews` | List access review schedule definitions |
| `list_conditional_access_policies` | List conditional access policies |
| `list_entitlement_access_packages` | List entitlement management access packages |
| `list_lifecycle_workflows` | List lifecycle management workflows |
| `update_conditional_access_policy` | Update a conditional access policy |

## Required Permissions
- `User.Invite.All, Policy.ReadWrite.ConditionalAccess, AccessReview.Read.All, EntitlementManagement.Read.All, LifecycleWorkflows.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
