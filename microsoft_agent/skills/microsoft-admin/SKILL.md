---
name: microsoft-admin
description: "Microsoft 365 Admin — Service Health, Messages, SharePoint Admin & Delegated Admin Relationships"
---

# Microsoft 365 Admin

Manage tenant administration including service health, announcements, SharePoint admin, and delegated admin relationships.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_admin_sharepoint` | Get SharePoint admin settings for the tenant |
| `get_delegated_admin_relationship` | Get a specific delegated admin relationship |
| `get_service_health` | Get the health status for a specific service |
| `get_service_health_issue` | Get a specific service health issue |
| `get_service_update_message` | Get a specific service update message |
| `list_delegated_admin_relationships` | List delegated admin relationships |
| `list_service_health` | Get the service health status for all services in the tenant |
| `list_service_health_issues` | List all service health issues for the tenant |
| `list_service_update_messages` | List service update messages (message center posts) for the tenant |
| `update_admin_sharepoint` | Update SharePoint admin settings for the tenant |

## Required Permissions
- `ServiceHealth.Read.All, ServiceMessage.Read.All, Sites.Read.All, DelegatedAdminRelationship.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
