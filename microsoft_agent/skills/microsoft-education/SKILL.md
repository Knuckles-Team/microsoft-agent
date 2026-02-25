---
name: microsoft-education
description: "Microsoft 365 Education — Education Classes, Schools, Users & Assignments"
tags: [education]
---

# Microsoft 365 Education

Manage education classes, schools, users, and assignments.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_education_class` | Get a specific education class |
| `get_education_school` | Get a specific education school |
| `list_education_assignments` | List assignments for an education class |
| `list_education_classes` | List education classes |
| `list_education_schools` | List education schools |
| `list_education_users` | List education users |

## Required Permissions
- `EduAdministration.ReadWrite`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
