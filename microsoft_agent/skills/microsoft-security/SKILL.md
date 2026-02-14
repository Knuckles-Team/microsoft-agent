---
name: microsoft-security
description: "Microsoft 365 Security — Alerts, Incidents, Threat Intelligence & Identity Protection"
---

# Microsoft 365 Security

Manage security alerts, incidents, secure scores, threat intelligence, identity protection, and information protection.

## Available Tools

| Tool | Description |
|------|-------------|
| `dismiss_risky_user` | Dismiss a risky user (mark as safe) |
| `get_risk_detection` | Get a specific risk detection |
| `get_risky_user` | Get a specific risky user |
| `get_security_alert` | Get a specific security alert by ID |
| `get_security_incident` | Get a specific security incident by ID |
| `get_sensitivity_label` | Get a specific sensitivity label |
| `get_threat_intelligence_host` | Get a specific threat intelligence host |
| `list_risk_detections` | List risk detections (sign-in anomalies, leaked credentials, etc.) |
| `list_risky_users` | List users flagged as risky |
| `list_secure_scores` | List tenant secure scores over time |
| `list_security_alerts` | List security alerts |
| `list_security_incidents` | List security incidents |
| `list_sensitivity_labels` | List sensitivity labels |
| `list_threat_intelligence_hosts` | List threat intelligence hosts |
| `run_hunting_query` | Run an advanced hunting query using Kusto Query Language (KQL) |
| `update_security_alert` | Update a security alert |
| `update_security_incident` | Update a security incident |

## Required Permissions
- `SecurityEvents.ReadWrite.All, SecurityIncident.ReadWrite.All, ThreatHunting.Read.All, IdentityRiskEvent.Read.All, IdentityRiskyUser.ReadWrite.All, InformationProtectionPolicy.Read`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
