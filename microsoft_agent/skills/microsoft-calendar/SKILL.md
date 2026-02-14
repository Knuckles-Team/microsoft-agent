---
name: microsoft-calendar
description: "Microsoft 365 Calendar — Calendar Events, Calendars & Scheduling"
---

# Microsoft 365 Calendar

Manage calendar events, calendars, scheduling, and meeting time suggestions.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_calendar_event` | TIP: CRITICAL: Do not try to guess the email address of the recipients |
| `create_specific_calendar_event` | TIP: CRITICAL: Do not try to guess the email address of the recipients |
| `delete_calendar_event` | delete_calendar_event: DELETE /me/events/{event-id} |
| `delete_specific_calendar_event` | delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id} |
| `find_meeting_times` | find_meeting_times: POST /me/findMeetingTimes |
| `get_calendar_event` | get_calendar_event: GET /me/events/{event-id} |
| `get_calendar_view` | get_calendar_view: GET /me/calendarView |
| `get_specific_calendar_event` | get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id} |
| `list_calendar_events` | list_calendar_events: GET /me/events |
| `list_calendars` | list_calendars: GET /me/calendars |
| `list_specific_calendar_events` | list_specific_calendar_events: GET /me/calendars/{calendar-id}/events |
| `update_calendar_event` | TIP: CRITICAL: Do not try to guess the email address of the recipients |
| `update_specific_calendar_event` | TIP: CRITICAL: Do not try to guess the email address of the recipients |

## Required Permissions
- `Calendars.ReadWrite`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
