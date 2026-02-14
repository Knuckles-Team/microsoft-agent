---
name: microsoft-search
description: "Microsoft 365 Search — Microsoft Graph Search Queries"
---

# Microsoft 365 Search

Execute Microsoft Graph search queries.

## Available Tools

| Tool | Description |
|------|-------------|
| `search_query` | search_query: POST /search/query |

## Required Permissions
- `Files.Read.All, Sites.Read.All, Mail.Read`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
