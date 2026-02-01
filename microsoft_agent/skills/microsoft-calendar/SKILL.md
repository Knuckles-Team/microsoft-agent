---
name: microsoft-calendar
description: "Generated skill for calendar operations. Contains 13 tools."
---

### Overview
This skill handles operations related to calendar.

### Available Tools
- `list_calendar_events`: list_calendar_events: GET /me/events
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
    - `timezone` (Optional[str])
- `get_calendar_event`: get_calendar_event: GET /me/events/{event-id}
  - **Parameters**:
    - `event_id` (str)
    - `params` (Optional[Dict[str, Any]])
    - `timezone` (Optional[str])
- `create_calendar_event`: create_calendar_event: POST /me/events  TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_calendar_event`: update_calendar_event: PATCH /me/events/{event-id}  TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
  - **Parameters**:
    - `event_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `delete_calendar_event`: delete_calendar_event: DELETE /me/events/{event-id}
  - **Parameters**:
    - `event_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_specific_calendar_events`: list_specific_calendar_events: GET /me/calendars/{calendar-id}/events
  - **Parameters**:
    - `calendar_id` (str)
    - `params` (Optional[Dict[str, Any]])
    - `timezone` (Optional[str])
- `get_specific_calendar_event`: get_specific_calendar_event: GET /me/calendars/{calendar-id}/events/{event-id}
  - **Parameters**:
    - `calendar_id` (str)
    - `event_id` (str)
    - `params` (Optional[Dict[str, Any]])
    - `timezone` (Optional[str])
- `create_specific_calendar_event`: create_specific_calendar_event: POST /me/calendars/{calendar-id}/events  TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
  - **Parameters**:
    - `calendar_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_specific_calendar_event`: update_specific_calendar_event: PATCH /me/calendars/{calendar-id}/events/{event-id}  TIP: CRITICAL: Do not try to guess the email address of the recipients. Use the list-users tool to find the email address of the recipients.
  - **Parameters**:
    - `calendar_id` (str)
    - `event_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `delete_specific_calendar_event`: delete_specific_calendar_event: DELETE /me/calendars/{calendar-id}/events/{event-id}
  - **Parameters**:
    - `calendar_id` (str)
    - `event_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_calendar_view`: get_calendar_view: GET /me/calendarView
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
    - `timezone` (Optional[str])
- `list_calendars`: list_calendars: GET /me/calendars
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `find_meeting_times`: find_meeting_times: POST /me/findMeetingTimes
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
