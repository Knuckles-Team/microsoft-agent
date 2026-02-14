---
name: microsoft-domains
description: "Microsoft 365 Domains — Tenant Domain Management & DNS Configuration"
---

# Microsoft 365 Domains

Manage tenant domains including adding, verifying, deleting, and viewing DNS configuration records.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_domain` | Add a domain to the tenant |
| `delete_domain` | Delete a domain from the tenant |
| `get_domain` | Get properties of a specific domain |
| `list_domain_service_configuration_records` | List DNS records required by the domain for Microsoft services |
| `list_domains` | List domains associated with the tenant |
| `verify_domain` | Verify ownership of a domain |

## Required Permissions
- `Domain.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
