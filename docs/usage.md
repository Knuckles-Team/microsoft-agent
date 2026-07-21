# Usage

Microsoft Agent exposes one governed Microsoft Graph client through an MCP server
and, when installed, an A2A agent. The MCP server is the normal integration point.
The Python client is available for trusted in-process extensions.

## MCP surface

`MCP_TOOL_MODE=intent` exposes the compact Agent Utilities intent and control
surface. It discovers the installed Microsoft action schemas at runtime and routes
requests while preserving caller identity, approval, policy, and trace context.
Use MCP discovery as the authority for exact schemas; do not copy generated action
lists into prompts or configuration.

The installed provider families cover:

- authentication, metadata, search, users, groups, and directory operations;
- mail, calendars, contacts, tasks, files, sites, notes, Teams, and chat;
- applications, policy, security, audit, reports, devices, storage, print, and
  subscriptions;
- keyed zero-PII projection and governed Epistemic Graph ingestion; and
- optional documents, Office bridge, Power Platform, Intune, and Windows
  companion capabilities.

Every routed action crosses the fail-closed Microsoft tool policy. Read actions are
available by default. Unknown actions are classified as writes. Non-destructive
writes require `MICROSOFT_ALLOW_WRITES=true`; destructive actions additionally
require `MICROSOFT_ALLOW_DESTRUCTIVE=true`.

`list_microsoft_ingestion_projection` returns only keyed opaque node identifiers,
types, and relationships. GraphOS validates the signed source manifest and
commits that projection through the native ChangeEnvelope seam. Projection requires a deployment-owned
pseudonymization key; neither returns or persists Microsoft content or identity
fields.

## MCP server

Use stdio when the MCP client owns the process:

```bash
microsoft-mcp --transport stdio
```

Use streamable HTTP only behind deployment-owned caller authentication and TLS:

```bash
microsoft-mcp --transport streamable-http --host 127.0.0.1 --port 8000
```

The supported transports are exactly `stdio` and `streamable-http`. A non-loopback
listener must not be exposed until its MCP caller-authentication policy, TLS trust,
and network controls are configured.

## Python client

`get_client()` returns the sole authenticated `MicrosoftGraphApi` authority. Its
methods are asynchronous.

```python
import asyncio

from microsoft_agent.auth import get_client


async def main() -> None:
    graph = await get_client()
    users = await graph.list_users(params={"$top": 25})
    events = await graph.list_calendar_events(params={"$top": 25})
    print(len(users.get("value", [])), len(events.get("value", [])))


asyncio.run(main())
```

Authentication settings are validated immediately before token acquisition. Do
not construct a second token client or persist access tokens outside the supplied
secure authentication boundary.

## A2A agent

Install the `agent` extra, then point the agent at an authenticated MCP endpoint or
an external MCP configuration:

```bash
microsoft-agent \
  --mcp-url <mcp-url> \
  --provider <provider> \
  --model-id <model-id>
```

Model endpoints, keys, TLS profiles, Langfuse settings, and skills remain external
Agent Utilities configuration. The package does not ship a provider, model, API
key, MCP URL, or local filesystem profile.

## Optional capability groups

Set `MICROSOFT_ENABLED_TOOL_GROUPS` to the explicit families approved for the
deployment. Optional document, Office bridge, Power Platform, Intune, and Windows
capabilities also require their corresponding package extras and external
integration configuration. Disabled groups are not registered.

See [Configuration](configuration.md) for identity and policy controls,
[Authentication](authentication.md) for supported token modes, and
[Deployment](deployment.md) for process and container examples.
