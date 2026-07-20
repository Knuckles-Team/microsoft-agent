# Microsoft Agent

Microsoft Agent provides a governed Microsoft Graph MCP server and an optional A2A
agent. It uses Agent Utilities for MCP, identity propagation, intent delegation,
configuration, telemetry, and graph integration.

Current package version: **1.0.1**

## Capabilities

- Outlook mail, calendars, contacts, tasks, and attachments
- Teams chats, channels, meetings, and presence
- OneDrive, SharePoint, Excel, OneNote, and Universal Print
- Entra users, groups, applications, policy, audit, security, and reports
- Optional Word and PowerPoint generation and an exact-origin Office.js bridge
- Optional Power Platform, Intune, and outbound Windows companion integrations
- Governed ingestion into Epistemic Graph with provider-owned ontology, mapping,
  schema fingerprint, provenance, and quarantine contracts
- One consolidated `microsoft-agent-operations` skill

The current implementation has one modular Graph client in
`microsoft_agent/api_client.py`, one authentication authority in
`microsoft_agent/auth.py`, and one MCP entry point in
`microsoft_agent/mcp_server.py`. No raw-token fallback, plaintext cache, bundled
database, bundled endpoint profile, or duplicate API wrapper is shipped.

## Install

```bash
python -m pip install "microsoft-agent[mcp]"
```

Useful extras:

```bash
python -m pip install "microsoft-agent[agent]"        # A2A agent runtime
python -m pip install "microsoft-agent[documents]"    # Word and PowerPoint
python -m pip install "microsoft-agent[cloud]"        # managed/workload identity
python -m pip install "microsoft-agent[windows]"      # Windows companion
python -m pip install "microsoft-agent[control-plane]"
python -m pip install "microsoft-agent[all]"
```

Agent Utilities depends on `epistemic-graph[full]`; the full and numeric engine
features therefore remain available when the A2A/GraphOS runtime is installed.

## Authentication

Microsoft identity values are supplied by deployment configuration. There is no
baked-in client identifier.

Supported modes:

- `delegated`: broker or browser login; device code is opt-in
- `application`: certificate or referenced client secret
- `on_behalf_of`: verified incoming user delegation
- `external_token`: verified request token for an explicitly allowed audience
- `managed_identity`: Azure managed identity
- `workload_identity`: federated workload token file

Token caches use OS-backed secure storage only. If secure storage is unavailable,
tokens remain in memory and are not written to plaintext files.

<!-- ENV-VARS-TABLE:START -->

| Variable | Purpose | Default |
|---|---|---|
| `MICROSOFT_TENANT_ID` | Entra tenant identifier | required except managed identity |
| `MICROSOFT_CLIENT_ID` | Registered application identifier | required except system-assigned managed identity |
| `MICROSOFT_AUTH_MODE` | Identity mode listed above | `delegated` |
| `MICROSOFT_LOGIN_METHOD` | `auto`, `broker`, `browser`, or enabled `device_code` | `auto` |
| `MICROSOFT_PERMISSION_PROFILES` | Named least-privilege permission bundles | `productivity,collaboration` |
| `MICROSOFT_GRAPH_SCOPES` | Additional deployment-approved Graph scopes | empty |
| `MICROSOFT_CLIENT_CERTIFICATE_PATH` | Deployment-owned certificate reference | empty |
| `MICROSOFT_CLIENT_CERTIFICATE_THUMBPRINT` | Certificate thumbprint | empty |
| `MICROSOFT_CLIENT_SECRET_REF` | `env://`, `vault://`, or `secret://` client-secret reference | empty |
| `MICROSOFT_MANAGED_IDENTITY_CLIENT_ID` | User-assigned managed identity | empty |
| `MICROSOFT_WORKLOAD_IDENTITY_TOKEN_FILE` | Federated workload token reference | empty |
| `MICROSOFT_ENABLED_TOOL_GROUPS` | Optional integration families | see `.env.example` |
| `MICROSOFT_ALLOW_WRITES` | Enable non-destructive writes | `false` |
| `MICROSOFT_ALLOW_DESTRUCTIVE` | Enable destructive operations when writes are enabled | `false` |
| `MICROSOFT_GRAPH_BASE_URL` | Configured Graph API root | Microsoft Graph v1.0 |
| `MICROSOFT_GRAPH_TLS_PROFILE` | Named Agent Utilities TLS profile for Graph | deployment default |
| `MICROSOFT_GRAPH_TLS_PROFILE_REF` | Secret reference to a Graph TLS profile | empty |
| `MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF` | Secret reference for the zero-PII graph identifier key | required for ingestion only |
| `MICROSOFT_INTEGRATIONS_CONFIG_PATH` | External integration and allowlist profile | empty |
| `MICROSOFT_OFFICE_ADDIN_ORIGINS` | Exact HTTPS Office add-in origins | empty |
| `MCP_TOOL_MODE` | Agent Utilities surface | `intent` |
| `TRANSPORT` | `stdio` or `streamable-http` | `stdio` |
| `AUTH_TYPE` | MCP caller authentication boundary | deployment-owned |

Credentials, endpoints, trust bundles, and integration profiles are never packaged
in the wheel.

<!-- ENV-VARS-TABLE:END -->

## Run the MCP server

Local stdio:

```bash
microsoft-mcp --transport stdio
```

Authenticated HTTP deployment:

```bash
microsoft-mcp --transport streamable-http --host 127.0.0.1 --port 8000
```

Do not expose an unauthenticated listener outside loopback. Microsoft Graph identity
and MCP caller identity are separate boundaries and both must be configured.

<!-- MCP-TOOLS-TABLE:START -->

`MCP_TOOL_MODE=intent` is the default small surface. It discovers and delegates to
the current Microsoft action tools while preserving caller authority and approval.
The provider action families include:

- `microsoft_auth`, `microsoft_meta`, and `microsoft_search`
- `microsoft_mail`, `microsoft_calendar`, `microsoft_files`, and `microsoft_notes`
- `microsoft_chat`, `microsoft_teams`, `microsoft_groups`, and `microsoft_user`
- administration, directory, policy, security, audit, reporting, device, print,
  storage, site, subscription, and application families
- `list_microsoft_ingestion_projection` for keyed zero-PII source projection
- signed GraphOS source sync for governed ChangeEnvelope materialization
- optional document, Office bridge, Power Platform, Intune, and Windows tools

Use MCP discovery for the exact installed schema. Do not copy stale per-operation
lists into prompts or configuration.

<!-- MCP-TOOLS-TABLE:END -->

Every routed action passes the fail-closed tool policy. Reads are permitted by
default; unknown actions are treated as writes. Writes require
`MICROSOFT_ALLOW_WRITES=true`, and destructive actions additionally require
`MICROSOFT_ALLOW_DESTRUCTIVE=true`.

## MCP client configuration

<!-- MCP-CONFIG-EXAMPLES:START -->

> **Install the connector-focused `[mcp]` extra.** Examples use `microsoft-agent[mcp]` to add
> FastMCP / FastAPI through `agent-utilities[mcp]`; the required Agent Utilities core
> still carries `epistemic-graph[full]`. The `[agent-runtime]` extra additionally
> enables model orchestration.

#### stdio Transport (local IDEs — Cursor, Claude Desktop, VS Code)

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": [
        "--from",
        "microsoft-agent[mcp]",
        "microsoft-mcp"
      ],
      "env": {
        "MCP_TOOL_MODE": "intent"
      }
    }
  }
}
```

Runtime references require an alias-aware launcher such as GraphOS. Other
launchers must omit those entries and inject the resolved values through their
own runtime secret boundary.

#### Streamable-HTTP Transport (networked / production)

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": [
        "--from",
        "microsoft-agent[mcp]",
        "microsoft-mcp",
        "--transport",
        "streamable-http",
        "--port",
        "8000"
      ],
      "env": {
        "TRANSPORT": "streamable-http",
        "HOST": "127.0.0.1",
        "PORT": "8000",
        "MCP_TOOL_MODE": "intent"
      }
    }
  }
}
```

Alternatively, connect to a pre-deployed Streamable-HTTP instance by `url`:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "url": "http://localhost:8000/microsoft-mcp/mcp"
    }
  }
}
```

Run a reviewed container image as a least-privilege stdio child (no
listener or published port):

```bash
docker run -i --rm \
  --read-only \
  --cap-drop=ALL \
  --security-opt=no-new-privileges \
  --pids-limit=256 \
  --tmpfs /tmp:rw,noexec,nosuid,nodev,size=64m \
  -e TRANSPORT=stdio \
  -e MCP_TOOL_MODE=intent \
  registry.example.invalid/microsoft-agent@sha256:<digest> microsoft-mcp
```

For containerized network HTTP, supply an authenticated TLS ingress (or
direct server TLS), exact `MCP_ALLOWED_HOSTS`, and an exact trusted-proxy
CIDR policy through the operator-owned deployment profile. The generator
does not emit an unauthenticated non-loopback listener.

_Auto-generated from the code-read env surface (`MCP_TOOL_MODE` + package vars) — do not edit._
<!-- MCP-CONFIG-EXAMPLES:END -->

## A2A agent

```bash
microsoft-agent --mcp-url <mcp-url> --provider <provider> --model-id <model-id>
```

The agent uses the Agent Utilities model/configuration boundary. Provider keys,
model endpoints, Langfuse configuration, and TLS profiles are supplied externally.

## Governed ingestion

The provider contributes:

- `microsoft_agent/connectors/mcp_source_presets.json`
- `connector_manifest.yml`
- `microsoft_agent/ontology/microsoft.ttl`
- SHACL, mapping, fixture, migration, schema-fingerprint, and certification assets

Microsoft source projection persists only keyed opaque identifiers, structural node
types, and relationships. It never stores names, addresses, subjects, bodies,
filenames, URLs, timestamps, attachment bytes, or provider identifiers. The
pseudonymization key is supplied by AgentConfig or a secret store and is never
packaged or traced. Records remain quarantined until tenant, ACL, provenance,
schema, signature, and privacy requirements are satisfied. Generated signatures
and fingerprints must be regenerated whenever the tool schema or ontology changes;
stale attestations are not valid release evidence.

## Optional integrations

Power Platform, Intune, Windows companion, document roots, Office origins, and
allowlists are described by an external profile referenced with
`MICROSOFT_INTEGRATIONS_CONFIG_PATH`. The repository contains only the native
connection points and generic schemas. It does not ship organization-specific
environments, device lists, URLs, or ontologies.

## Containers

`docker/Dockerfile` has `mcp` and `agent` targets:

```bash
docker build -f docker/Dockerfile --target mcp -t <registry>/microsoft-agent:<version>-mcp .
docker build -f docker/Dockerfile --target agent -t <registry>/microsoft-agent:<version> .
```

Deployment-owned Compose or Kubernetes configuration supplies identity, trust,
storage, and network policy.

## Documentation

- [Overview](docs/overview.md)
- [Installation](docs/installation.md)
- [Authentication](docs/authentication.md)
- [Configuration](docs/configuration.md)
- [Usage](docs/usage.md)
- [Deployment](docs/deployment.md)
- [Office add-in](docs/office-addin.md)
- [Windows companion](docs/windows-companion.md)
- [Integration enrollment checklist](docs/enrollment-checklist.md)

The MkDocs navigation is defined in `mkdocs.yml`; strict documentation builds are a
release gate.

## Development

```bash
uv sync --all-extras --dev
uv run pytest
uv run ruff check .
uv run python scripts/security_sanitizer.py
uv run python scripts/security_contract.py --contract .security/security-contract.json validate
uv run python scripts/verify_api_integration.py --local
uv run mkdocs build --strict
```

See `AGENTS.md` for current-only, privacy, security, and change-discipline rules.


<!-- BEGIN agent-utilities-deployment (generated; do not edit between markers) -->

## Deploy with `agent-utilities-deployment`

Provision this package with the consolidated **`agent-utilities-deployment`**
workflow. It selects an installed-package, editable-source, or immutable-container
path; records only runtime secret and TLS-profile references in `AgentConfig`; and
runs doctor, registration, policy, observability, and rollback gates. Ask your agent
to **"deploy `microsoft-agent` with agent-utilities-deployment"**.

| Install mode | Command |
|------|---------|
| Installed package | `uv tool install "microsoft-agent[mcp]"`, then run `microsoft-mcp` |
| Editable source | `uv pip install -e ".[agent]"`, then run `microsoft-mcp` |
| Immutable container | deploy `registry.example.invalid/microsoft-agent@sha256:<digest>` through the operator-selected orchestrator |

The repository embeds no deployment profile, credential value, certificate path, or
environment-specific endpoint. Supply those at runtime through `AgentConfig` and the
configured secret provider.

<!-- END agent-utilities-deployment -->
