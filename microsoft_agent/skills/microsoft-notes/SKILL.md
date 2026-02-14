---
name: microsoft-notes
description: "Microsoft 365 Notes — OneNote Notebooks, Sections & Pages"
---

# Microsoft 365 Notes

Manage OneNote notebooks, sections, and pages.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_onenote_page` | create_onenote_page: POST /me/onenote/pages |
| `get_onenote_page_content` | get_onenote_page_content: GET /me/onenote/pages/{onenotePage-id}/content |
| `list_onenote_notebook_sections` | list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections |
| `list_onenote_notebooks` | list_onenote_notebooks: GET /me/onenote/notebooks |
| `list_onenote_section_pages` | list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages |

## Required Permissions
- `Notes.ReadWrite`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
