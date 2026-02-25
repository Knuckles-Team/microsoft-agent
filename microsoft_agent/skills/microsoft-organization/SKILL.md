---
name: microsoft-organization
description: "Microsoft 365 Organization — Organization Profile, Branding & Configuration"
tags: [organization]
---

# Microsoft 365 Organization

Manage organization profile, branding, and configuration.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_org_branding` | Get organization branding properties (sign-in page customization) |
| `get_organization` | Get a specific organization by ID |
| `list_organization` | Get the properties and relationships of the currently authenticated organization |
| `update_org_branding` | Update organization branding properties |
| `update_organization` | Update organization properties |

## Required Permissions
- `Organization.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
