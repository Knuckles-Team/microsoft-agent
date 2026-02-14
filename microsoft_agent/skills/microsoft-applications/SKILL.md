---
name: microsoft-applications
description: "Microsoft 365 Applications — App Registrations, Service Principals & Enterprise Apps"
---

# Microsoft 365 Applications

Manage app registrations, service principals, credentials, and enterprise apps.

## Available Tools

| Tool | Description |
|------|-------------|
| `add_application_password` | Add a password credential (client secret) to an app |
| `create_application` | Create an app registration |
| `create_service_principal` | Create a service principal for an app |
| `delete_application` | Delete an app registration |
| `delete_service_principal` | Delete a service principal |
| `get_application` | Get a specific app registration |
| `get_service_principal` | Get a specific service principal |
| `list_applications` | List app registrations in the tenant |
| `list_service_principals` | List service principals (enterprise apps) |
| `remove_application_password` | Remove a password credential from an app |
| `update_application` | Update an app registration |
| `update_service_principal` | Update a service principal |

## Required Permissions
- `Application.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
