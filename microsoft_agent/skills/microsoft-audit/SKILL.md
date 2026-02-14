---
name: microsoft-audit
description: "Microsoft 365 Audit — Directory Audits, Sign-In Logs & Provisioning Logs"
---

# Microsoft 365 Audit

Access directory audit logs, sign-in logs, and provisioning logs.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_directory_audit` | Get a specific directory audit entry |
| `get_sign_in_log` | Get a specific sign-in log entry |
| `list_directory_audits` | List directory audit log entries |
| `list_provisioning_logs` | List provisioning logs |
| `list_sign_in_logs` | List sign-in activity logs |

## Required Permissions
- `AuditLog.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
