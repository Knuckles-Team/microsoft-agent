---
name: microsoft-sites
description: "Microsoft 365 Sites — SharePoint Sites, Lists, Drives & Items"
---

# Microsoft 365 Sites

Manage SharePoint sites, site lists, site drives, site items, and site administration.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_admin_sharepoint` | Get SharePoint admin settings for the tenant |
| `get_sharepoint_site_by_path` | get_sharepoint_site_by_path: GET /sites/{hostname}:/{server-relative-path} |
| `get_sharepoint_site_list_item` | get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id} |
| `get_sharepoint_sites_delta` | get_sharepoint_sites_delta: GET /sites/delta() |
| `get_site` | get_site: GET /sites/{site-id} |
| `get_site_drive_by_id` | get_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id} |
| `get_site_item` | get_site_item: GET /sites/{site-id}/items/{baseItem-id} |
| `get_site_list` | Get a specific SharePoint site list |
| `list_sharepoint_site_list_items` | List items in a SharePoint site list |
| `list_site_drives` | list_site_drives: GET /sites/{site-id}/drives |
| `list_site_items` | list_site_items: GET /sites/{site-id}/items |
| `list_site_lists` | List lists for a SharePoint site |
| `list_sites` | list_sites: GET /sites |
| `update_admin_sharepoint` | Update SharePoint admin settings for the tenant |

## Required Permissions
- `Sites.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
