---
name: microsoft-subscriptions
description: "Microsoft 365 Subscriptions — Webhook Subscriptions for Change Notifications"
---

# Microsoft 365 Subscriptions

Manage webhook subscriptions for change notifications.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_subscription` | Create a webhook subscription for change notifications |
| `delete_subscription` | Delete a webhook subscription |
| `get_subscription` | Get a specific subscription |
| `list_subscriptions` | List active webhook subscriptions for change notifications |
| `update_subscription` | Renew a subscription by extending its expiration time |

## Required Permissions
- `Subscription varies by resource`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
