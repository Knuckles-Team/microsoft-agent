---
name: microsoft-document-authoring
skill_type: skill
description: >-
  Generate Word documents and PowerPoint presentations, upload them to OneDrive,
  and drive open Word and PowerPoint sessions through the Office add-in bridge.
  Use when the agent must produce or edit an Office document. Do NOT use for
  spreadsheet ranges (use microsoft-office-apps).
license: MIT
tags: [microsoft, word, powerpoint, office]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Word and PowerPoint authoring

Document generation and Office add-in bridge operations.

## When to use
- Generate a document or deck and upload it.
- Pair with an Office session; read or write the Word selection.
- Add, list or delete PowerPoint slides.

## When NOT to use
- Excel: use microsoft-office-apps.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Steps

1. Check what can be generated with `get_document_capabilities`.
   Verify the needed document type is listed.
2. Generate the file with `generate_word_document` or `generate_powerpoint_presentation`, or with the `generate_and_upload_*` variant to place it in OneDrive.
   Verify the response names the file or its destination path.
3. For a live Office session, call `list_office_sessions` and use the returned session_id with `get_word_selection_from_office` or `list_powerpoint_slides_in_office`.
   Verify the session id is present before editing.
4. Edit through the session with `write_word_selection_in_office` or `add_powerpoint_slide_in_office` and poll `get_office_command_result`.
   Verify the command result reports success.

## Failure recovery

| Failure | Cause | Recovery |
|---|---|---|
| Writes refused | `MICROSOFT_ALLOW_WRITES` is not set on the server | Set `MICROSOFT_ALLOW_WRITES` in the server environment, restart the server, and submit the call with the same arguments |
| No Office session listed | The add-in is not paired or not open | Call `create_office_pairing`, open the add-in with the returned pairing, and confirm `list_office_sessions` shows a session |
| Destructive call refused | No live confirmation was given | Obtain the gateway approval or pass `confirm` set to true in the call, as the Safety section requires |

## Tools (15)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `add_powerpoint_slide_in_office` | W | session_id, wait_timeout_seconds |
| `create_office_pairing` | W | expected_host, label |
| `delete_powerpoint_slide_in_office` | D | session_id, slide_number, wait_timeout_seconds |
| `generate_and_upload_powerpoint_presentation` | W | conflict_behavior, destination_path, drive_id, request |
| `generate_and_upload_word_document` | W | conflict_behavior, destination_path, drive_id, request |
| `generate_powerpoint_presentation` | W | request |
| `generate_word_document` | W | request |
| `get_document_capabilities` | R | none |
| `get_office_command_result` | R | command_id |
| `get_word_selection_from_office` | R | session_id, wait_timeout_seconds |
| `insert_powerpoint_text_box_in_office` | W | height, left, session_id, slide_number, text, top, wait_timeout_seconds, width |
| `list_office_sessions` | R | none |
| `list_powerpoint_slides_in_office` | R | session_id, wait_timeout_seconds |
| `replace_word_placeholders_in_office` | W | placeholder, replacement, session_id, wait_timeout_seconds |
| `write_word_selection_in_office` | W | content, mode, session_id, wait_timeout_seconds |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `documents-powerpoint-slide`: `list_powerpoint_slides_in_office` -> `add_powerpoint_slide_in_office` -> `delete_powerpoint_slide_in_office`
- `documents-word-selection`: `get_word_selection_from_office` -> `write_word_selection_in_office`
- `documents-read-records`: `get_document_capabilities` -> `get_office_command_result` -> `list_office_sessions`
- `documents-change-records`: `create_office_pairing` -> `generate_and_upload_powerpoint_presentation` -> `generate_and_upload_word_document` -> `generate_powerpoint_presentation` -> `generate_word_document` -> `insert_powerpoint_text_box_in_office` -> `replace_word_placeholders_in_office`
