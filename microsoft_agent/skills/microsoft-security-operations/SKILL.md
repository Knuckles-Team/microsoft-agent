---
name: microsoft-security-operations
skill_type: skill
description: >-
  Security alerts and incidents, risky users and risk detections, sign-in and
  directory audit logs, secure score, threat intelligence, sensitivity labels,
  subject rights requests and advanced hunting on Microsoft Graph. Use when the
  agent must triage security signals or read audit evidence. Do NOT use for
  access administration (use microsoft-identity-directory).
license: MIT
tags: [microsoft, security, audit, graph]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft security operations

Typed security operations. Updating an alert or incident changes its state: record the reason.

## When to use
- Triage alerts and incidents; update their status.
- List risky users and detections; dismiss a risky user.
- Read sign-in, directory audit and provisioning logs.
- Run an advanced hunting query; read secure scores.

## When NOT to use
- Changing roles or policies: use microsoft-identity-directory.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (24)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_subject_rights_request` | W | data, params |
| `microsoft_dismiss_risky_user` | W | params, user_id |
| `microsoft_get_directory_audit` | R | audit_id, params |
| `microsoft_get_risk_detection` | R | params, risk_id |
| `microsoft_get_risky_user` | R | params, user_id |
| `microsoft_get_security_alert` | R | alert_id, params |
| `microsoft_get_security_incident` | R | incident_id, params |
| `microsoft_get_sensitivity_label` | R | label_id, params |
| `microsoft_get_sign_in_log` | R | params, sign_in_id |
| `microsoft_get_subject_rights_request` | R | params, request_id |
| `microsoft_get_threat_intelligence_host` | R | host_id, params |
| `microsoft_list_directory_audits` | R | params |
| `microsoft_list_risk_detections` | R | params |
| `microsoft_list_risky_users` | R | params |
| `microsoft_list_secure_scores` | R | params |
| `microsoft_list_security_alerts` | R | params |
| `microsoft_list_security_incidents` | R | params |
| `microsoft_list_sensitivity_labels` | R | params |
| `microsoft_list_sign_in_logs` | R | params |
| `microsoft_list_subject_rights_requests` | R | params |
| `microsoft_list_threat_intelligence_hosts` | R | params |
| `microsoft_run_hunting_query` | W | data, params |
| `microsoft_update_security_alert` | W | alert_id, data, params |
| `microsoft_update_security_incident` | W | data, incident_id, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `security-directory-audit`: `microsoft_get_directory_audit` -> `microsoft_list_directory_audits`
- `security-risk-detection`: `microsoft_get_risk_detection` -> `microsoft_list_risk_detections`
- `security-risky-user`: `microsoft_get_risky_user` -> `microsoft_list_risky_users` -> `microsoft_dismiss_risky_user`
- `security-alert`: `microsoft_get_security_alert` -> `microsoft_list_security_alerts` -> `microsoft_update_security_alert`
- `security-incident`: `microsoft_get_security_incident` -> `microsoft_list_security_incidents` -> `microsoft_update_security_incident`
- `security-sensitivity-label`: `microsoft_get_sensitivity_label` -> `microsoft_list_sensitivity_labels`
- `security-sign-log`: `microsoft_get_sign_in_log` -> `microsoft_list_sign_in_logs`
- `security-subject-right`: `microsoft_get_subject_rights_request` -> `microsoft_list_subject_rights_requests` -> `microsoft_create_subject_rights_request`
- `security-threat-intelligence`: `microsoft_get_threat_intelligence_host` -> `microsoft_list_threat_intelligence_hosts`
- `security-read-records`: `microsoft_list_secure_scores`
- `security-change-records`: `microsoft_run_hunting_query`
