---
name: microsoft-calendar-scheduling
skill_type: skill
description: >-
  Outlook calendar and Teams online-meeting scheduling on Microsoft Graph via the
  microsoft-agent MCP server — list a calendar view, read/create/update/delete
  events, and find meeting times with the domain-typed `microsoft_calendar` tool.
  Use when the agent must check availability, book or reschedule a meeting, or
  enumerate a user's events in a date window. Do NOT use for mailbox messages
  (use microsoft-mail-operations) or Teams chat posts
  (use microsoft-teams-messaging).
license: MIT
tags: [microsoft, outlook, calendar, meetings, mcp]
metadata:
  author: Genius
  version: '0.1.0'
---
# Microsoft Calendar & Scheduling

Domain-typed access to Outlook calendars and events through Microsoft Graph
(`/me/events`, `/me/calendarView`, `/me/findMeetingTimes`). Prefer the
`microsoft_calendar` tool over raw Graph calls — it carries the event and
meeting-time action set and returns Graph-shaped records.

## When to use
- Enumerate events in a window (`get_calendar_view` with start/end).
- Read / create / update / delete a calendar event.
- Suggest slots that work for a set of attendees (`find_meeting_times`).
- List the user's calendars or events on a specific calendar.

## When NOT to use
- Mailbox triage / sending mail → `microsoft-mail-operations`.
- Teams chat or channel messages → `microsoft-teams-messaging`.
- Files / documents → the `microsoft_files` tool.

## Prerequisites & environment
Connect via the `mcp-client` skill against the **`microsoft-agent`** MCP server.
Run `microsoft_auth` (`login`) first. Requires the `Calendars.ReadWrite` (and,
for meeting links, `OnlineMeetings.ReadWrite`) Graph scopes.

| Variable | Required | Notes |
|----------|----------|-------|
| `OIDC_CLIENT_ID` | optional | Azure AD app (client) id; a default is baked in |
| `TESTING` | optional | Set truthy to skip the global auth manager |

## Tools & actions
Prefer the **condensed** `microsoft_calendar` tool; `action` + a `params_json`
**JSON string** (query options inside a `params` object).

| Condensed tool | Key actions |
|----------------|-------------|
| `microsoft_calendar` | `list_calendar_events`, `get_calendar_view`, `get_calendar_event`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event`, `list_calendars`, `find_meeting_times` |

## Recipes (`params_json`)
List a calendar view for a date window (expanded recurrences):
```json
{"params":{"startDateTime":"2026-07-01T00:00:00Z","endDateTime":"2026-07-08T00:00:00Z","$orderby":"start/dateTime","$top":50}}
```
Get one event:
```json
{"event_id":"AAMkAG...","params":{"$select":"subject,start,end,organizer,attendees"}}
```
Create an event (body carries the event payload):
```json
{"body":{"subject":"Design sync","start":{"dateTime":"2026-07-05T15:00:00","timeZone":"UTC"},"end":{"dateTime":"2026-07-05T15:30:00","timeZone":"UTC"},"attendees":[{"emailAddress":{"address":"ana@contoso.com"},"type":"required"}]}}
```

## Gotchas
- `params_json` is a **string** of JSON — serialize it.
- `get_calendar_view` requires **both** `startDateTime` and `endDateTime` and is
  the only way to get expanded recurring instances; plain `list_calendar_events`
  returns the series masters.
- Times are objects: `{"dateTime": ..., "timeZone": ...}` — pass a `timeZone` or
  Graph assumes UTC.
- Responses are raw Graph JSON — the collection is under `value`.

## Related
- **Native KG ingestion:** the `microsoft_ingest_records` tool (`kind:"events"`)
  maps events → `:Event` (+ organizer/attendee `:Person`) nodes.
- **Source preset:** `microsoft-calendar` in `connectors/mcp_source_presets.json`
  syncs events into the KG as `event` documents.
