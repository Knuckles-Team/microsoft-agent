---
name: microsoft-teams
description: "Microsoft 365 Teams — Teams, Channels, Messages & Membership"
---

# Microsoft 365 Teams

Manage teams, channels, channel messages, and team membership.

## Available Tools

| Tool | Description |
|------|-------------|
| `get_channel_message` | get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id} |
| `get_team` | get_team: GET /teams/{team-id} |
| `get_team_channel` | get_team_channel: GET /teams/{team-id}/channels/{channel-id} |
| `list_channel_messages` | list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages |
| `list_joined_teams` | list_joined_teams: GET /me/joinedTeams |
| `list_team_channels` | list_team_channels: GET /teams/{team-id}/channels |
| `list_team_members` | list_team_members: GET /teams/{team-id}/members |
| `send_channel_message` | send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages |

## Required Permissions
- `Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Read.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
