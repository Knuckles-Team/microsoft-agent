---
name: microsoft-groups
description: "Microsoft 365 Groups — Groups, Membership, Ownership & Group Conversations"
---

# Microsoft 365 Groups

Manage Microsoft 365 groups, security groups, membership, ownership, and group conversations.

## Available Tools

| Tool | Description |
|------|-------------|
| `add_group_member` | Add a member to a group |
| `create_group` | Create a new Microsoft 365 group or security group |
| `delete_group` | Delete a group |
| `get_group` | Get properties and relationships of a group object |
| `list_group_conversations` | List conversations in a Microsoft 365 group |
| `list_group_drives` | List drives (document libraries) of a group |
| `list_group_members` | Get a list of the group |
| `list_group_owners` | Get owners of a group |
| `list_groups` | List all Microsoft 365 groups and security groups in the organization |
| `remove_group_member` | Remove a member from a group |
| `update_group` | Update properties of a group object |

## Required Permissions
- `Group.ReadWrite.All, GroupMember.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
