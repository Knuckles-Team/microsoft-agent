---
name: microsoft-tasks
description: "Generated skill for tasks operations. Contains 13 tools."
---

### Overview
This skill handles operations related to tasks.

### Available Tools
- `list_todo_task_lists`: list_todo_task_lists: GET /me/todo/lists
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `list_todo_tasks`: list_todo_tasks: GET /me/todo/lists/{todoTaskList-id}/tasks
  - **Parameters**:
    - `todoTaskList_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_todo_task`: get_todo_task: GET /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}
  - **Parameters**:
    - `todoTaskList_id` (str)
    - `todoTask_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `create_todo_task`: create_todo_task: POST /me/todo/lists/{todoTaskList-id}/tasks
  - **Parameters**:
    - `todoTaskList_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_todo_task`: update_todo_task: PATCH /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}
  - **Parameters**:
    - `todoTaskList_id` (str)
    - `todoTask_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `delete_todo_task`: delete_todo_task: DELETE /me/todo/lists/{todoTaskList-id}/tasks/{todoTask-id}
  - **Parameters**:
    - `todoTaskList_id` (str)
    - `todoTask_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_planner_tasks`: list_planner_tasks: GET /me/planner/tasks
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `get_planner_plan`: get_planner_plan: GET /planner/plans/{plannerPlan-id}
  - **Parameters**:
    - `plannerPlan_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_plan_tasks`: list_plan_tasks: GET /planner/plans/{plannerPlan-id}/tasks
  - **Parameters**:
    - `plannerPlan_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_planner_task`: get_planner_task: GET /planner/tasks/{plannerTask-id}
  - **Parameters**:
    - `plannerTask_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `create_planner_task`: create_planner_task: POST /planner/tasks
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_planner_task`: update_planner_task: PATCH /planner/tasks/{plannerTask-id}
  - **Parameters**:
    - `plannerTask_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_planner_task_details`: update_planner_task_details: PATCH /planner/tasks/{plannerTask-id}/details
  - **Parameters**:
    - `plannerTask_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
