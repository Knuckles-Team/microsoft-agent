---
name: microsoft-files-sharepoint
skill_type: skill
description: >-
  OneDrive drives, SharePoint sites, lists and list items, folder contents, file
  upload and download, and SharePoint admin settings on Microsoft Graph. Use
  when the agent must find, read, upload or delete a file or read a SharePoint
  list. Do NOT use for Teams messages (use microsoft-teams-messaging) or Office
  document content (use microsoft-office-apps).
license: MIT
tags: [microsoft, onedrive, sharepoint, files, graph]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft OneDrive and SharePoint files

Typed drive and site operations. Download returns file content; upload writes it.

## When to use
- Browse drives, sites, lists and folders.
- Download or upload a file; delete a OneDrive file.
- Read SharePoint list items and site deltas.
- Read file storage containers and storage reports.

## When NOT to use
- Spreadsheet or notebook content: use microsoft-office-apps.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Steps

1. Locate the site or drive with `microsoft_list_sites` or `microsoft_list_drives`.
   Verify the site name or drive id matches the request.
2. Browse contents with `microsoft_list_folder_files` or `microsoft_list_site_lists`.
   Verify the item id you need appears in the result.
3. Read data with `microsoft_download_onedrive_file_content` or `microsoft_list_sharepoint_site_list_items`.
   Verify the file name or item count matches expectations.
4. Change data only when asked with `microsoft_upload_file_content` or `microsoft_delete_onedrive_file`.
   Verify the item appears or disappears in `microsoft_list_folder_files`.

## Failure recovery

| Failure | Cause | Recovery |
|---|---|---|
| Authentication error or expired session | The Graph session is missing or expired | Call `microsoft_login`, call `microsoft_verify_login`, and resubmit the original call (see microsoft-session-management) |
| Writes refused | `MICROSOFT_ALLOW_WRITES` is not set on the server | Set `MICROSOFT_ALLOW_WRITES` in the server environment, restart the server, and submit the call with the same arguments |
| Destructive call refused | No live confirmation was given | Obtain the gateway approval or pass `confirm` set to true in the call, as the Safety section requires |

## Tools (26)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_file_storage_container` | W | data, params |
| `microsoft_delete_onedrive_file` | D | confirm, driveItem_id, drive_id, params |
| `microsoft_download_onedrive_file_content` | W | driveItem_id, drive_id, params |
| `microsoft_get_admin_sharepoint` | R | params |
| `microsoft_get_drive_root_item` | R | drive_id, params |
| `microsoft_get_file_storage_container` | R | container_id, params |
| `microsoft_get_onedrive_usage_report` | R | params, period |
| `microsoft_get_sharepoint_activity_report` | R | params, period |
| `microsoft_get_sharepoint_site_by_path` | R | params, path, site_id |
| `microsoft_get_sharepoint_site_list_item` | R | listItem_id, list_id, params, site_id |
| `microsoft_get_sharepoint_sites_delta` | R | continuation_url, fetch_all, max_pages, params |
| `microsoft_get_site` | R | params, site_id |
| `microsoft_get_site_drive_by_id` | R | drive_id, params, site_id |
| `microsoft_get_site_item` | R | baseItem_id, params, site_id |
| `microsoft_get_site_list` | R | list_id, params, site_id |
| `microsoft_list_drives` | R | params |
| `microsoft_list_file_storage_containers` | R | params |
| `microsoft_list_folder_files` | R | driveItem_id, drive_id, params |
| `microsoft_list_group_drives` | R | group_id, params |
| `microsoft_list_sharepoint_site_list_items` | R | list_id, params, site_id |
| `microsoft_list_site_drives` | R | params, site_id |
| `microsoft_list_site_items` | R | params, site_id |
| `microsoft_list_site_lists` | R | params, site_id |
| `microsoft_list_sites` | R | params |
| `microsoft_update_admin_sharepoint` | W | data, params |
| `microsoft_upload_file_content` | W | data, driveItem_id, drive_id, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `files-admin-sharepoint`: `microsoft_get_admin_sharepoint` -> `microsoft_update_admin_sharepoint`
- `files-file-storage`: `microsoft_get_file_storage_container` -> `microsoft_list_file_storage_containers` -> `microsoft_create_file_storage_container`
- `files-onedrive-file`: `microsoft_download_onedrive_file_content` -> `microsoft_delete_onedrive_file`
- `files-sharepoint-site`: `microsoft_get_sharepoint_site_by_path` -> `microsoft_get_sharepoint_site_list_item` -> `microsoft_get_sharepoint_sites_delta` -> `microsoft_list_sharepoint_site_list_items`
- `files-site`: `microsoft_get_site` -> `microsoft_list_sites`
- `files-site-drive`: `microsoft_get_site_drive_by_id` -> `microsoft_list_site_drives`
- `files-site-item`: `microsoft_get_site_item` -> `microsoft_list_site_items`
- `files-site-list`: `microsoft_get_site_list` -> `microsoft_list_site_lists`
- `files-read-records`: `microsoft_get_drive_root_item` -> `microsoft_get_onedrive_usage_report` -> `microsoft_get_sharepoint_activity_report` -> `microsoft_list_drives` -> `microsoft_list_folder_files` -> `microsoft_list_group_drives`
- `files-change-records`: `microsoft_upload_file_content`
