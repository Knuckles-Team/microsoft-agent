---
name: microsoft-service-reports
skill_type: skill
description: >-
  Microsoft 365 service health, health issues, service update messages, and
  usage reports for mailbox, OneDrive, SharePoint, Teams and active users. Use
  when the agent must report tenant health or adoption. Do NOT use for security
  events (use microsoft-security-operations).
license: MIT
tags: [microsoft, reports, health]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft service health and usage reports

Read-only reporting operations.

## When to use
- Report current service health and open issues.
- Read message-center updates.
- Pull usage and activity reports.

## When NOT to use
- Security events: use microsoft-security-operations.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Steps

1. List current health with `microsoft_list_service_health`.
   Verify each service shows a status.
2. Drill into problems with `microsoft_list_service_health_issues` and `microsoft_get_service_health_issue`.
   Verify the issue id and start time.
3. Read announcements with `microsoft_list_service_update_messages`.
   Verify the message category matches the request.
4. Report adoption with `microsoft_get_office365_active_users` and a period such as D30.
   Verify the report returns rows for that period.

## Failure recovery

| Failure | Cause | Recovery |
|---|---|---|
| Authentication error or expired session | The Graph session is missing or expired | Call `microsoft_login`, call `microsoft_verify_login`, and resubmit the original call (see microsoft-session-management) |
| Report call fails | The period value is not one the report accepts | Pass one of the period values named in the tool description, such as D7 or D30, and resubmit |
| Empty or unexpected result | The id or filter does not match any object | List the parent collection with the list tool in the step above, copy the id from the result, and resubmit |

## Tools (7)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_get_office365_active_users` | R | params, period |
| `microsoft_get_service_health` | R | params, service_name |
| `microsoft_get_service_health_issue` | R | issue_id, params |
| `microsoft_get_service_update_message` | R | message_id, params |
| `microsoft_list_service_health` | R | params |
| `microsoft_list_service_health_issues` | R | params |
| `microsoft_list_service_update_messages` | R | params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `reports-service-health`: `microsoft_get_service_health` -> `microsoft_get_service_health_issue` -> `microsoft_list_service_health` -> `microsoft_list_service_health_issues`
- `reports-service-update`: `microsoft_get_service_update_message` -> `microsoft_list_service_update_messages`
- `reports-read-records`: `microsoft_get_office365_active_users`
