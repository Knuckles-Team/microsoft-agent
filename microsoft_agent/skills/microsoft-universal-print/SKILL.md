---
name: microsoft-universal-print
skill_type: skill
description: >-
  Universal Print printers, shares, print jobs and document upload sessions. Use
  when the agent must list printers or submit a print job. Do NOT use for file
  storage.
license: MIT
tags: [microsoft, print]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Universal Print

Printing creates a job on a real printer: confirm the printer and document.

## When to use
- List printers and shares.
- Create a job, upload a document, start the job.

## When NOT to use
- Document storage: use microsoft-files-sharepoint.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (8)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_print_document_upload_session` | W | content_type, document_name, print_document_id, print_job_id, printer_id, size |
| `microsoft_create_print_job` | W | data, params, printer_id |
| `microsoft_get_printer` | R | params, printer_id |
| `microsoft_list_print_jobs` | R | params, printer_id |
| `microsoft_list_print_shares` | R | params |
| `microsoft_list_printers` | R | params |
| `microsoft_start_print_job` | W | print_job_id, printer_id |
| `microsoft_submit_print_document` | R | printer_id, submission, upload_transport |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `print-document`: `microsoft_create_print_document_upload_session` -> `microsoft_submit_print_document`
- `print-job`: `microsoft_list_print_jobs` -> `microsoft_create_print_job` -> `microsoft_start_print_job`
- `printer`: `microsoft_get_printer` -> `microsoft_list_printers`
- `print-read-records`: `microsoft_list_print_shares`
