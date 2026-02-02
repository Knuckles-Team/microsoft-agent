---
name: microsoft-sites
description: "Generated skill for sites operations. Contains 12 tools."
---

### Overview
This skill handles operations related to sites.

### Available Tools
- `search_sharepoint_sites`: search_sharepoint_sites: GET /sites
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site`: get_sharepoint_site: GET /sites/{site-id}
  - **Parameters**:
    - `site_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_sharepoint_site_drives`: list_sharepoint_site_drives: GET /sites/{site-id}/drives
  - **Parameters**:
    - `site_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site_drive_by_id`: get_sharepoint_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id}
  - **Parameters**:
    - `site_id` (str)
    - `drive_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_sharepoint_site_items`: list_sharepoint_site_items: GET /sites/{site-id}/items
  - **Parameters**:
    - `site_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site_item`: get_sharepoint_site_item: GET /sites/{site-id}/items/{baseItem-id}
  - **Parameters**:
    - `site_id` (str)
    - `baseItem_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_sharepoint_site_lists`: list_sharepoint_site_lists: GET /sites/{site-id}/lists
  - **Parameters**:
    - `site_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site_list`: get_sharepoint_site_list: GET /sites/{site-id}/lists/{list-id}
  - **Parameters**:
    - `site_id` (str)
    - `list_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_sharepoint_site_list_items`: list_sharepoint_site_list_items: GET /sites/{site-id}/lists/{list-id}/items
  - **Parameters**:
    - `site_id` (str)
    - `list_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site_list_item`: get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id}
  - **Parameters**:
    - `site_id` (str)
    - `list_id` (str)
    - `listItem_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_site_by_path`: get_sharepoint_site_by_path: GET /sites/{site-id}/getByPath(path='{path}')
  - **Parameters**:
    - `site_id` (str)
    - `path` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_sharepoint_sites_delta`: get_sharepoint_sites_delta: GET /sites/delta()
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
