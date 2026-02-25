---
name: microsoft-communications
description: "Microsoft 365 Communications — Online Meetings, Call Records & Presence"
tags: [communications]
---

# Microsoft 365 Communications

Manage online meetings, call records, and user presence information.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_online_meeting` | Create a new online meeting |
| `delete_online_meeting` | Delete an online meeting |
| `get_call_record` | Get a specific call record by ID |
| `get_my_presence` | Get current user |
| `get_online_meeting` | Get a specific online meeting by ID |
| `get_presence` | Get presence for a specific user by user ID |
| `list_call_records` | List call records |
| `list_online_meetings` | List online meetings for the current user |
| `list_presences` | List presence information for multiple users |
| `update_online_meeting` | Update an existing online meeting |

## Required Permissions
- `OnlineMeetings.ReadWrite, CallRecords.Read.All, Presence.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
