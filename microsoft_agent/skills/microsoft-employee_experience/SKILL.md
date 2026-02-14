---
name: microsoft-employee_experience
description: "Microsoft 365 Employee Experience — Learning Providers & Course Activities"
---

# Microsoft 365 Employee Experience

Manage learning providers and course activities for employee development.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_learning_provider` | Get a specific learning provider |
| `list_learning_course_activities` | List learning course activities for the current user |
| `list_learning_providers` | List learning providers |

## Required Permissions
- `LearningContent.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
