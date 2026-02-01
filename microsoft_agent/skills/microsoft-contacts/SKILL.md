---
name: microsoft-contacts
description: "Generated skill for contacts operations. Contains 5 tools."
---

### Overview
This skill handles operations related to contacts.

### Available Tools
- `list_outlook_contacts`: list_outlook_contacts: GET /me/contacts
  - **Parameters**:
    - `params` (Optional[Dict[str, Any]])
- `get_outlook_contact`: get_outlook_contact: GET /me/contacts/{contact-id}
  - **Parameters**:
    - `contact_id` (str)
    - `params` (Optional[Dict[str, Any]])
- `create_outlook_contact`: create_outlook_contact: POST /me/contacts
  - **Parameters**:
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `update_outlook_contact`: update_outlook_contact: PATCH /me/contacts/{contact-id}
  - **Parameters**:
    - `contact_id` (str)
    - `data` (Optional[Dict[str, Any]])
    - `params` (Optional[Dict[str, Any]])
- `delete_outlook_contact`: delete_outlook_contact: DELETE /me/contacts/{contact-id}
  - **Parameters**:
    - `contact_id` (str)
    - `params` (Optional[Dict[str, Any]])

### Usage Instructions
1. Review the tool available in this skill.
2. Call the tool with the required parameters.

### Error Handling
- Ensure all required parameters are provided.
- Check return values for error messages.
