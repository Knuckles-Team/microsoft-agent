---
name: microsoft-directory
description: "Microsoft 365 Directory — Directory Objects, Roles, Role Definitions & Role Assignments"
tags: [directory]
---

# Microsoft 365 Directory

Manage directory objects, roles, deleted items, role definitions, and role assignments.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_role_assignment` | Create a new RBAC role assignment |
| `get_directory_object` | Get a specific directory object |
| `get_directory_role` | Get a specific activated directory role |
| `get_role_assignment` | Get a specific RBAC role assignment |
| `get_role_definition` | Get a specific RBAC role definition |
| `list_deleted_items` | List recently deleted directory items (users, groups, apps) |
| `list_directory_objects` | List directory objects |
| `list_directory_role_templates` | List all directory role templates (built-in role definitions) |
| `list_directory_roles` | List activated directory roles |
| `list_role_assignments` | List RBAC directory role assignments |
| `list_role_definitions` | List RBAC directory role definitions |
| `restore_deleted_item` | Restore a recently deleted directory item |

## Required Permissions
- `Directory.ReadWrite.All, RoleManagement.ReadWrite.Directory`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
