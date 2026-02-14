---
name: microsoft-connections
description: "Microsoft 365 Connections — Microsoft Search External Connections"
---

# Microsoft 365 Connections

Manage Microsoft Search external connections for custom data ingestion.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_external_connection` | Create an external connection for Microsoft Search |
| `delete_external_connection` | Delete an external connection |
| `get_external_connection` | Get a specific external connection |
| `list_external_connections` | List Microsoft Search external connections |

## Required Permissions
- `ExternalConnection.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
