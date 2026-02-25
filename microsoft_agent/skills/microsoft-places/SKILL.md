---
name: microsoft-places
description: "Microsoft 365 Places — Rooms & Room Lists"
tags: [places]
---

# Microsoft 365 Places

Manage rooms and room lists.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_place` | Get a specific place (room or room list) |
| `list_room_lists` | List room lists |
| `list_rooms` | List conference rooms |
| `update_place` | Update a place (room) |

## Required Permissions
- `Place.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
