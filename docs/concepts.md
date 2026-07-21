# Capability concepts

> **Prefix**: `CONCEPT:MSFT-*`
> **Version**: 0.15.0
> **Bridge**: [`CONCEPT:AU-ECO.messaging.native-backend-abstraction`](https://github.com/Knuckles-Team/agent-utilities/blob/main/docs/concepts.md) (Unified Toolkit Ingestion)
Microsoft Agent contributes Microsoft Graph provider semantics to the ecosystem; it
does not replace Agent Utilities or Epistemic Graph authorities.

| Concept | Provider responsibility | Ecosystem authority |
|---|---|---|
| Microsoft identity | acquire a tenant- and audience-bound Graph token through one validated mode | Agent Utilities verifies MCP caller and delegation identity |
| Graph actions | execute bounded asynchronous operations through one modular client | Agent Utilities owns tool discovery, session, approval, and trace context |
| Action policy | classify the routed action and fail closed for writes/destructive actions | permission governance supplies caller policy and approval |
| Optional integrations | expose native document, Power Platform, Intune, Office, and Windows connection points | deployment AgentConfig supplies allowlists, endpoints, trust, and secrets |
| Knowledge ingestion | map approved provider records into quarantined change envelopes | Epistemic Graph owns persistence, ACL, provenance, lineage, and deletion semantics |
| Microsoft skill | provide one consolidated operational workflow and provider catalog | GraphOS owns skill discovery and direct/delegated execution |
| Observability | attach privacy-safe provider outcome metadata | Agent Utilities and Langfuse own trace transport and retention policy |

## Provider domains

| Concept ID | Name | Description |
|------------|------|-------------|
| `CONCEPT:MS-OS.governance.msft` | Administration | MCP tool domain `admin` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-2` | Agreements Operations | MCP tool domain `agreements` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-3` | Applications Operations | MCP tool domain `applications` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-4` | Audit Operations | MCP tool domain `audit` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-5` | Authentication & Session Management | MCP tool domain `auth` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-6` | Calendar Management | MCP tool domain `calendar` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-7` | Chat & Messaging | MCP tool domain `chat` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-8` | Communications Operations | MCP tool domain `communications` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-9` | Connections Operations | MCP tool domain `connections` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-10` | Contact Management | MCP tool domain `contacts` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-11` | Devices Operations | MCP tool domain `devices` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-12` | Directory Operations | MCP tool domain `directory` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-13` | Domains Operations | MCP tool domain `domains` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-14` | Education Operations | MCP tool domain `education` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-15` | Employee Experience Operations | MCP tool domain `employee_experience` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-16` | File Management | MCP tool domain `files` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-17` | Group Management | MCP tool domain `groups` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-18` | Identity Operations | MCP tool domain `identity` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-19` | Email & Messaging | MCP tool domain `mail` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-20` | Meta Operations | MCP tool domain `meta` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-21` | Notes Operations | MCP tool domain `notes` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-22` | Organization Operations | MCP tool domain `organization` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-23` | Places Operations | MCP tool domain `places` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-24` | Policies Operations | MCP tool domain `policies` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-25` | Print Operations | MCP tool domain `print` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-26` | Privacy Operations | MCP tool domain `privacy` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-27` | Reports Operations | MCP tool domain `reports` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-28` | Search & Discovery | MCP tool domain `search` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-29` | Security Operations | MCP tool domain `security` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-30` | Sites Operations | MCP tool domain `sites` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-31` | Solutions Operations | MCP tool domain `solutions` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-32` | Storage & Persistence | MCP tool domain `storage` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-33` | Subscriptions Operations | MCP tool domain `subscriptions` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-34` | Tasks Operations | MCP tool domain `tasks` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-35` | Teams Operations | MCP tool domain `teams` — Action-routed dynamic tool registration |
| `CONCEPT:MS-OS.governance.msft-36` | User & Identity Management | MCP tool domain `user` — Action-routed dynamic tool registration |
The action catalog covers authentication and metadata; users, groups, directory,
applications, policies, audit, security, and reports; Outlook, calendar, contacts,
tasks, files, sites, Excel, OneNote, Teams, and chat; and supporting device,
storage, print, communications, subscription, and search domains.

The installed MCP schema is the authoritative domain/action inventory. Documentation
describes capability families only so it cannot silently diverge from dynamically
registered tool schemas.

| Concept ID | Name | Origin |
|------------|------|--------|
| `CONCEPT:AU-ECO.messaging.native-backend-abstraction` | Unified Toolkit Ingestion | agent-utilities |
| `CONCEPT:AU-ORCH.adapter.hot-cache-invalidation` | Confidence-Gated Router | agent-utilities |
| `CONCEPT:AU-OS.config.secrets-authentication` | Prompt Injection Defense | agent-utilities |
| `CONCEPT:AU-OS.state.cognitive-scheduler-preemption` | Cognitive Scheduler | agent-utilities |
| `CONCEPT:AU-OS.governance.reactive-multi-axis-budget` | Guardrail Engine | agent-utilities |
| `CONCEPT:AU-OS.governance.wasm-micro-agent-sandbox` | Audit Logging | agent-utilities |
| `CONCEPT:AU-KG.query.object-graph-mapper` | Knowledge Graph Core | agent-utilities |

## Synergy with agent-utilities

This project integrates with `agent-utilities` via `CONCEPT:AU-ECO.messaging.native-backend-abstraction` (Unified Toolkit Ingestion). The `microsoft_agent` MCP server registers its tools with the agent-utilities FastMCP middleware, enabling automatic discovery, telemetry, and Knowledge Graph ingestion of all MSFT-* concepts.
## Invariants

- One Graph client and one Microsoft authentication manager exist.
- Microsoft identity never substitutes for MCP caller identity.
- Unknown actions fail closed as writes.
- Optional integrations remain absent until dependencies and external configuration
  are both ready.
- Provider records remain quarantined until schema, tenant, ACL, provenance,
  signature, retention, and privacy checks pass.
- Generated fingerprints and certification metadata are invalid after any tool or
  ontology change and must be regenerated before release.
