---
name: microsoft-auth
description: "Generated skill for auth operations. Contains 4 tools."
---

### Overview
This skill handles operations related to auth.

### Available Tools
- `login`: Authenticate with Microsoft using device code flow
  - **Parameters**:
    - `force` (bool)
- `logout`: Log out from Microsoft account
- `verify_login`: Check current Microsoft authentication status
- `list_accounts`: List all available Microsoft accounts

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
