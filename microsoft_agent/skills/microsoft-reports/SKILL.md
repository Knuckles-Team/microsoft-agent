---
name: microsoft-reports
description: "Microsoft 365 Reports — Usage & Activity Reports"
---

# Microsoft 365 Reports

Generate usage and activity reports for email, mailbox, Office 365, SharePoint, Teams, and OneDrive.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_email_activity_report` | Get email activity user detail report |
| `get_mailbox_usage_report` | Get mailbox usage detail report |
| `get_office365_active_users` | Get Office 365 active user detail report |
| `get_onedrive_usage_report` | Get OneDrive usage account detail report |
| `get_sharepoint_activity_report` | Get SharePoint activity user detail report |
| `get_teams_user_activity` | Get Teams user activity detail report |

## Required Permissions
- `Reports.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
