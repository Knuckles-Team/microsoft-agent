---
name: microsoft-teams
description: "Generated skill for teams operations. Contains 8 tools."
---

### Overview
This skill handles operations related to teams.

### Available Tools
- `list_joined_teams`: list_joined_teams: GET /me/joinedTeams
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `get_team`: get_team: GET /teams/{team-id}
  - **Parameters**:
    - `team_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_team_channels`: list_team_channels: GET /teams/{team-id}/channels
  - **Parameters**:
    - `team_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_team_channel`: get_team_channel: GET /teams/{team-id}/channels/{channel-id}
  - **Parameters**:
    - `team_id` (str)
    - `channel_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `list_channel_messages`: list_channel_messages: GET /teams/{team-id}/channels/{channel-id}/messages
  - **Parameters**:
    - `team_id` (str)
    - `channel_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `get_channel_message`: get_channel_message: GET /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}
  - **Parameters**:
    - `team_id` (str)
    - `channel_id` (str)
    - `chatMessage_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `send_channel_message`: send_channel_message: POST /teams/{team-id}/channels/{channel-id}/messages
  - **Parameters**:
    - `team_id` (str)
    - `channel_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `list_team_members`: list_team_members: GET /teams/{team-id}/members
  - **Parameters**:
    - `team_id` (str)
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
