---
name: microsoft-files
description: "Microsoft 365 Files — OneDrive, Excel, OneNote & SharePoint Files"
---

# Microsoft 365 Files

Manage OneDrive files, Excel workbooks, OneNote notebooks, and SharePoint file operations.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_excel_chart` | create_excel_chart: POST /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/charts/add |
| `delete_onedrive_file` | delete_onedrive_file: DELETE /drives/{drive-id}/items/{driveItem-id} |
| `download_onedrive_file_content` | download_onedrive_file_content: GET /drives/{drive-id}/items/{driveItem-id}/content |
| `format_excel_range` | format_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/for |
| `get_drive_root_item` | get_drive_root_item: GET /drives/{drive-id}/root |
| `get_excel_table` | get_excel_table: GET /drives/{drive-id}/items/{item-id}/workbook/tables/{table-id} |
| `get_excel_workbook` | get_excel_workbook: GET /drives/{drive-id}/items/{item-id}/workbook |
| `get_excel_worksheet` | get_excel_worksheet: GET /drives/{drive-id}/items/{item-id}/workbook/worksheets/{worksheet-id} |
| `get_excel_range` | get_excel_range: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{worksheet-id}/range() |
| `get_root_folder` | get_root_folder: GET /drives/{drive-id}/root |
| `get_sharepoint_site_list_item` | get_sharepoint_site_list_item: GET /sites/{site-id}/lists/{list-id}/items/{listItem-id} |
| `get_site_drive_by_id` | get_site_drive_by_id: GET /sites/{site-id}/drives/{drive-id} |
| `get_site_item` | get_site_item: GET /sites/{site-id}/items/{baseItem-id} |
| `get_site_list` | Get a specific SharePoint site list |
| `list_calendar_events` | list_calendar_events: GET /me/events |
| `list_calendars` | list_calendars: GET /me/calendars |
| `list_channel_messages` | list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages |
| `list_chat_message_replies` | list_chat_message_replies: GET /chats/{chat-id}/messages/{chatMessage-id}/replies |
| `list_chat_messages` | list_chat_messages: GET /chats/{chat-id}/messages |
| `list_chats` | list_chats: GET /me/chats |
| `list_drives` | list_drives: GET /me/drives |
| `list_excel_tables` | List Excel tables in a workbook |
| `list_excel_worksheets` | list_excel_worksheets: GET /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets |
| `list_folder_files` | list_folder_files: GET /drives/{drive-id}/items/{driveItem-id}/children |
| `list_group_drives` | List drives (document libraries) of a group |
| `list_joined_teams` | list_joined_teams: GET /me/joinedTeams |
| `list_mail_attachments` | list_mail_attachments: GET /me/messages/{message-id}/attachments |
| `list_mail_folder_messages` | list_mail_folder_messages: GET /me/mailFolders/{mailFolder-id}/messages |
| `list_mail_messages` | list_mail_messages: GET /me/messages |
| `list_mail_folders` | list_mail_folders: GET /me/mailFolders |
| `list_onenote_notebook_sections` | list_onenote_notebook_sections: GET /me/onenote/notebooks/{notebook-id}/sections |
| `list_onenote_notebooks` | list_onenote_notebooks: GET /me/onenote/notebooks |
| `list_onenote_section_pages` | list_onenote_section_pages: GET /me/onenote/sections/{onenoteSection-id}/pages |
| `list_outlook_contacts` | list_outlook_contacts: GET /me/contacts |
| `list_shared_mailbox_folder_messages` | list_shared_mailbox_folder_messages: GET /users/{user-id}/mailFolders/{mailFolder-id}/messages |
| `list_shared_mailbox_messages` | list_shared_mailbox_messages: GET /users/{user-id}/messages |
| `list_plan_tasks` | list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks |
| `list_planner_tasks` | list_planner_tasks: GET /me/planner/tasks |
| `list_sharepoint_site_list_items` | List items in a SharePoint site list |
| `list_site_drives` | list_site_drives: GET /sites/{site-id}/drives |
| `list_site_items` | list_site_items: GET /sites/{site-id}/items |
| `list_site_lists` | List lists for a SharePoint site |
| `list_specific_calendar_events` | list_specific_calendar_events: GET /me/calendars/{calendar-id}/events |
| `list_team_channels` | list_team_channels: GET /teams/{team-id}/channels |
| `list_team_members` | list_team_members: GET /teams/{team-id}/members |
| `list_todo_task_lists` | list_todo_task_lists: GET /me/todo/lists |
| `list_todo_tasks` | list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks |
| `list_users` | list_users: GET /users |
| `sort_excel_range` | sort_excel_range: PATCH /drives/{drive-id}/items/{driveItem-id}/workbook/worksheets/{workbookWorksheet-id}/range()/sort |
| `upload_file_content` | upload_file_content: PUT /drives/{drive-id}/items/{driveItem-id}/content |

## Required Permissions
- `Files.ReadWrite, Sites.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
