---
name: microsoft-storage
description: "Microsoft 365 Storage — File Storage Containers"
tags: [storage]
---

# Microsoft 365 Storage

Manage file storage containers.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_file_storage_container` | Create a file storage container |
| `get_file_storage_container` | Get a specific file storage container |
| `list_file_storage_containers` | List file storage containers |

## Required Permissions
- `FileStorageContainer.Selected`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
