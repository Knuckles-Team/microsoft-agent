---
name: microsoft-tasks
description: "Microsoft 365 Tasks — Planner Tasks & To-Do Task Lists"
tags: [tasks]
---

# Microsoft 365 Tasks

Manage Planner tasks, plans, To-Do task lists, and task operations.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_planner_task` | create_planner_task: POST /planner/tasks |
| `create_todo_task` | create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks |
| `delete_todo_task` | delete_todo_task: DELETE /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id} |
| `get_planner_plan` | get_planner_plan: GET /planner/plans/{plannerPlan-id} |
| `get_planner_task` | get_planner_task: GET /planner/tasks/{plannerTask-id} |
| `get_todo_task` | get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id} |
| `list_plan_tasks` | list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks |
| `list_planner_tasks` | list_planner_tasks: GET /me/planner/tasks |
| `list_todo_task_lists` | list_todo_task_lists: GET /me/todo/lists |
| `list_todo_tasks` | list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks |
| `update_planner_task` | update_planner_task: PATCH /planner/tasks/{plannerTask-id} |
| `update_planner_task_details` | update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details |
| `update_todo_task` | update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id} |

## Required Permissions
- `Tasks.ReadWrite, Group.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
