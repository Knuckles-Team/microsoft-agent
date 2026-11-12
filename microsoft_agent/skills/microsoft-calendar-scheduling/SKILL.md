---
name: microsoft-calendar-scheduling
skill_type: skill
description: >-
  Outlook calendars, events, online meetings, meeting-time search, rooms and
  places on Microsoft Graph. Use when the agent must read a schedule, create or
  change an event, find a common meeting time, or manage an online meeting. Do
  NOT use for mail (use microsoft-mail-operations) or Teams channel messages
  (use microsoft-teams-messaging).
license: MIT
tags: [microsoft, calendar, outlook, graph, mcp]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft calendar scheduling

Typed calendar operations including the calendar view, specific-calendar events, online meetings and `find_meeting_times`.

## When to use
- Read the calendar view for a time window.
- Create, update or cancel an event.
- Find meeting times for several attendees.
- Create or delete an online meeting; list rooms and places.

## When NOT to use
- Mail: use microsoft-mail-operations.
- Contacts: use microsoft-contacts.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (23)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_calendar_event` | W | data, params |
| `microsoft_create_online_meeting` | W | data, params |
| `microsoft_create_specific_calendar_event` | W | calendar_id, data, params |
| `microsoft_delete_calendar_event` | D | confirm, event_id, params |
| `microsoft_delete_online_meeting` | D | confirm, meeting_id, params |
| `microsoft_delete_specific_calendar_event` | D | calendar_id, confirm, event_id, params |
| `microsoft_find_meeting_times` | R | data, params |
| `microsoft_get_calendar_event` | R | event_id, params, timezone |
| `microsoft_get_calendar_view` | R | params, timezone |
| `microsoft_get_online_meeting` | R | meeting_id, params |
| `microsoft_get_place` | R | params, place_id |
| `microsoft_get_specific_calendar_event` | R | calendar_id, event_id, params, timezone |
| `microsoft_list_calendar_events` | R | params, timezone |
| `microsoft_list_calendars` | R | params |
| `microsoft_list_online_meetings` | R | params |
| `microsoft_list_room_lists` | R | params |
| `microsoft_list_rooms` | R | params |
| `microsoft_list_specific_calendar_events` | R | calendar_id, params, timezone |
| `microsoft_list_virtual_events` | R | params |
| `microsoft_update_calendar_event` | W | data, event_id, params |
| `microsoft_update_online_meeting` | W | data, meeting_id, params |
| `microsoft_update_place` | W | data, params, place_id |
| `microsoft_update_specific_calendar_event` | W | calendar_id, data, event_id, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `calendar-event`: `microsoft_get_calendar_event` -> `microsoft_list_calendar_events` -> `microsoft_create_calendar_event` -> `microsoft_update_calendar_event` -> `microsoft_delete_calendar_event`
- `calendar-online-meeting`: `microsoft_get_online_meeting` -> `microsoft_list_online_meetings` -> `microsoft_create_online_meeting` -> `microsoft_update_online_meeting` -> `microsoft_delete_online_meeting`
- `calendar-place`: `microsoft_get_place` -> `microsoft_update_place`
- `calendar-specific-calendar`: `microsoft_get_specific_calendar_event` -> `microsoft_list_specific_calendar_events` -> `microsoft_create_specific_calendar_event` -> `microsoft_update_specific_calendar_event` -> `microsoft_delete_specific_calendar_event`
- `calendar-read-records`: `microsoft_find_meeting_times` -> `microsoft_get_calendar_view` -> `microsoft_list_calendars` -> `microsoft_list_room_lists` -> `microsoft_list_rooms` -> `microsoft_list_virtual_events`
