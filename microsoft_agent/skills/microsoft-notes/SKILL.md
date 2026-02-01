---
name: microsoft-notes
description: "Generated skill for notes operations. Contains 5 tools."
---

### Overview
This skill handles operations related to notes.

### Available Tools
- `list_onenote_notebooks`: list_onenote_notebooks: GET /me/onenote/notebooks
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `list_onenote_notebook_sections`: list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections
  - **Parameters**:
    - `notebook_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_onenote_section_pages`: list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages
  - **Parameters**:
    - `onenoteSection_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_onenote_page_content`: get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content
  - **Parameters**:
    - `onenotePage_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `create_onenote_page`: create_onenote_page: POST /me/onenote/pages
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
