---
name: microsoft-privacy
description: "Microsoft 365 Privacy — Subject Rights Requests (GDPR/CCPA)"
---

# Microsoft 365 Privacy

Manage subject rights requests for GDPR/CCPA compliance.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_subject_rights_request` | Create a subject rights request |
| `get_subject_rights_request` | Get a specific subject rights request |
| `list_subject_rights_requests` | List subject rights requests (GDPR/CCPA) |

## Required Permissions
- `SubjectRightsRequest.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
