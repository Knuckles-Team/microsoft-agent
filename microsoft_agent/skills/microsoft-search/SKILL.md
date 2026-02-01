---
name: microsoft-search
description: "Generated skill for search operations. Contains 2 tools."
---

### Overview
This skill handles operations related to search.

### Available Tools
- `search_sharepoint_sites`: search_sharepoint_sites: GET /sites
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `search_query`: search_query: POST /search/query
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
