---
name: microsoft-print
description: "Microsoft 365 Print — Printers, Print Jobs & Print Shares"
---

# Microsoft 365 Print

Manage printers, print jobs, and print shares.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_print_job` | Create a print job |
| `get_printer` | Get a specific printer |
| `list_print_jobs` | List print jobs for a printer |
| `list_print_shares` | List printer shares |
| `list_printers` | List printers registered in the tenant |

## Required Permissions
- `PrintJob.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
