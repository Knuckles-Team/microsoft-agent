# Usage — API / CLI / MCP

`microsoft-agent` exposes the same capability three ways: as **MCP tools** an agent
calls, as a **Python API** (`MicrosoftGraphApi`) you import, and as **command-line
servers**. The complete tool surface and the agent architecture are described in
[Overview](overview.md).

## As an MCP server

Once [deployed](deployment.md), the server registers domain-routed tools across 36
Microsoft Graph domains. Each domain is enabled or disabled with its `*TOOL`
environment flag, so the registered surface can be scoped per deployment.

| Group | Domains |
|---|---|
| People & identity | `user`, `groups`, `directory`, `identity`, `contacts`, `organization` |
| Productivity | `calendar`, `mail`, `files`, `notes`, `tasks`, `chat`, `teams` |
| Administration | `admin`, `applications`, `devices`, `domains`, `policies`, `security`, `reports`, `audit` |
| Platform | `search`, `sites`, `storage`, `solutions`, `subscriptions`, `print`, `places`, `education` |

Tools follow the `action_resource` naming convention (for example `list_user`,
`get_group`, `send_mail`, `post_events`). Example agent prompts that map onto these
tools:

- *"List the members of the 'Engineering' group."* → `list_members_group`
- *"Send a status email to the operations distribution list."* → `send_mail`
- *"Schedule a meeting with the Engineering team next Tuesday."* → `post_events`

## As a Python API

`MicrosoftGraphApi` is a layered client over the Microsoft Graph SDK, organized by
domain (mail, calendar, drive, directory, applications, administration). Build one
from the environment with the `get_client()` helper:

```python
import asyncio
from microsoft_agent.auth import get_client

async def main():
    api = await get_client()          # reads MICROSOFT_* from the environment / .env

    # Reads
    users = api.list_user()
    group = api.get_group(group_id="<group-id>")
    events = api.list_calendarview()

asyncio.run(main())
```

Construct the client directly from an `AuthManager` when you manage authentication
yourself:

```python
from microsoft_agent.auth import AuthManager, AUTHORITY, SCOPES
from microsoft_agent.api_client import MicrosoftGraphApi

auth = AuthManager(client_id="<client-id>", authority=AUTHORITY, scopes=SCOPES)
api = MicrosoftGraphApi(auth)
```

### Writes

Write operations follow the same `action_resource` convention and require an
application granted the corresponding Microsoft Graph permissions:

```python
api.send_mail(message={"subject": "Status", "body": {"content": "All green."}})
api.post_events(event={"subject": "Sync", "start": {...}, "end": {...}})
```

## As CLI servers

Two console scripts are installed.

### MCP server CLI (`microsoft-mcp`)

| Short | Long | Description |
|---|---|---|
| `-t` | `--transport` | `stdio`, `http`, or `sse` (default: `stdio`) |
| `-s` | `--host` | Host address for HTTP transport (default: `0.0.0.0`) |
| `-p` | `--port` | Port for HTTP transport (default: `8000`) |
|  | `--auth-type` | `none`, `static`, `jwt`, `oauth-proxy`, `oidc-proxy` |

```bash
microsoft-mcp --transport http --host 0.0.0.0 --port 8000
```

### A2A agent CLI (`microsoft-agent`)

| Argument | Description | Default |
|---|---|---|
| `--host` | Host to bind the agent server to | `0.0.0.0` |
| `--port` | Port to bind the agent server to | `9000` |
| `--provider` | LLM provider (`openai`, `anthropic`, `google`, `huggingface`) | `openai` |
| `--model-id` | LLM model id | `nvidia/nemotron-3-super` |
| `--mcp-url` | MCP server URL the agent consumes | `http://microsoft-agent:8000/mcp` |

```bash
microsoft-agent --provider openai --model-id gpt-4o --api-key sk-... \
  --mcp-url http://localhost:8000/mcp
```

See [Deployment](deployment.md) for running both servers together under Docker
Compose.
