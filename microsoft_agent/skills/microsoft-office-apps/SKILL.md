---
name: microsoft-office-apps
skill_type: skill
description: >-
  Excel workbooks, OneNote notebooks and pages, Planner plans and tasks, and
  Microsoft To Do lists on Microsoft Graph. Use when the agent must read or edit
  a workbook range, read or create a OneNote page, or manage planner and to-do
  tasks. Do NOT use for file storage (use microsoft-files-sharepoint) or Word
  and PowerPoint generation (use microsoft-document-authoring).
license: MIT
tags: [microsoft, excel, onenote, planner, todo]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Excel, OneNote, Planner and To Do

Typed operations over workbook ranges and tables, notebook pages, plan tasks and to-do tasks.

## When to use
- Read or format an Excel range, table or chart.
- Read notebooks, sections and pages; create a page.
- Create, update or complete planner and to-do tasks.

## When NOT to use
- Word and PowerPoint documents: use microsoft-document-authoring.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (26)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_excel_chart` | W | data, drive_id, item_id, params, worksheet_id |
| `microsoft_create_onenote_page` | W | data, params |
| `microsoft_create_planner_task` | W | data, params |
| `microsoft_create_todo_task` | W | data, params, todoTaskList_id |
| `microsoft_delete_todo_task` | D | confirm, params, todoTaskList_id, todoTask_id |
| `microsoft_format_excel_range` | W | address, data, drive_id, item_id, params, worksheet_id |
| `microsoft_get_excel_range` | R | address, drive_id, item_id, params, worksheet_id |
| `microsoft_get_excel_table` | R | drive_id, item_id, params, table_id |
| `microsoft_get_excel_workbook` | R | drive_id, item_id, params |
| `microsoft_get_excel_worksheet` | R | drive_id, item_id, params, worksheet_id |
| `microsoft_get_onenote_page_content` | R | onenotePage_id, params |
| `microsoft_get_planner_plan` | R | params, plannerPlan_id |
| `microsoft_get_planner_task` | R | params, plannerTask_id |
| `microsoft_get_todo_task` | R | params, todoTaskList_id, todoTask_id |
| `microsoft_list_excel_tables` | R | drive_id, item_id, params |
| `microsoft_list_excel_worksheets` | R | drive_id, item_id, params |
| `microsoft_list_onenote_notebook_sections` | R | notebook_id, params |
| `microsoft_list_onenote_notebooks` | R | params |
| `microsoft_list_onenote_section_pages` | R | onenoteSection_id, params |
| `microsoft_list_planner_tasks` | R | params |
| `microsoft_list_todo_task_lists` | R | params |
| `microsoft_list_todo_tasks` | R | params, todoTaskList_id |
| `microsoft_sort_excel_range` | W | address, data, drive_id, item_id, params, worksheet_id |
| `microsoft_update_planner_task` | W | data, etag, params, plannerTask_id |
| `microsoft_update_planner_task_details` | W | data, etag, params, plannerTask_id |
| `microsoft_update_todo_task` | W | data, params, todoTaskList_id, todoTask_id |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `office-apps-excel-range`: `microsoft_get_excel_range` -> `microsoft_format_excel_range` -> `microsoft_sort_excel_range`
- `office-apps-excel-table`: `microsoft_get_excel_table` -> `microsoft_list_excel_tables`
- `office-apps-excel-worksheet`: `microsoft_get_excel_worksheet` -> `microsoft_list_excel_worksheets`
- `office-apps-onenote-notebook`: `microsoft_list_onenote_notebook_sections` -> `microsoft_list_onenote_notebooks`
- `office-apps-onenote-page`: `microsoft_get_onenote_page_content` -> `microsoft_create_onenote_page`
- `office-apps-planner-task`: `microsoft_get_planner_task` -> `microsoft_list_planner_tasks` -> `microsoft_create_planner_task` -> `microsoft_update_planner_task` -> `microsoft_update_planner_task_details`
- `office-apps-todo-task`: `microsoft_get_todo_task` -> `microsoft_list_todo_task_lists` -> `microsoft_list_todo_tasks` -> `microsoft_create_todo_task` -> `microsoft_update_todo_task` -> `microsoft_delete_todo_task`
- `office-apps-read-records`: `microsoft_get_excel_workbook` -> `microsoft_get_planner_plan` -> `microsoft_list_onenote_section_pages`
- `office-apps-change-records`: `microsoft_create_excel_chart`
