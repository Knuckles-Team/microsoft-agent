# Concept Registry — microsoft-agent

> **Prefix**: `CONCEPT:MSFT-*`
> **Version**: 0.15.0
> **Bridge**: [`CONCEPT:ECO-4.0`](../../agent-utilities/docs/concepts.md) (Unified Toolkit Ingestion)

---

## Project-Specific Concepts

| Concept ID | Name | Description |
|------------|------|-------------|
| `CONCEPT:MSFT-001` | Administration | MCP tool domain `admin` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-002` | Agreements Operations | MCP tool domain `agreements` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-003` | Applications Operations | MCP tool domain `applications` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-004` | Audit Operations | MCP tool domain `audit` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-005` | Authentication & Session Management | MCP tool domain `auth` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-006` | Calendar Management | MCP tool domain `calendar` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-007` | Chat & Messaging | MCP tool domain `chat` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-008` | Communications Operations | MCP tool domain `communications` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-009` | Connections Operations | MCP tool domain `connections` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-010` | Contact Management | MCP tool domain `contacts` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-011` | Devices Operations | MCP tool domain `devices` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-012` | Directory Operations | MCP tool domain `directory` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-013` | Domains Operations | MCP tool domain `domains` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-014` | Education Operations | MCP tool domain `education` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-015` | Employee Experience Operations | MCP tool domain `employee_experience` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-016` | File Management | MCP tool domain `files` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-017` | Group Management | MCP tool domain `groups` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-018` | Identity Operations | MCP tool domain `identity` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-019` | Email & Messaging | MCP tool domain `mail` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-020` | Meta Operations | MCP tool domain `meta` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-021` | Notes Operations | MCP tool domain `notes` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-022` | Organization Operations | MCP tool domain `organization` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-023` | Places Operations | MCP tool domain `places` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-024` | Policies Operations | MCP tool domain `policies` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-025` | Print Operations | MCP tool domain `print` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-026` | Privacy Operations | MCP tool domain `privacy` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-027` | Reports Operations | MCP tool domain `reports` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-028` | Search & Discovery | MCP tool domain `search` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-029` | Security Operations | MCP tool domain `security` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-030` | Sites Operations | MCP tool domain `sites` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-031` | Solutions Operations | MCP tool domain `solutions` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-032` | Storage & Persistence | MCP tool domain `storage` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-033` | Subscriptions Operations | MCP tool domain `subscriptions` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-034` | Tasks Operations | MCP tool domain `tasks` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-035` | Teams Operations | MCP tool domain `teams` — Action-routed dynamic tool registration |
| `CONCEPT:MSFT-036` | User & Identity Management | MCP tool domain `user` — Action-routed dynamic tool registration |

## Cross-Project References (from agent-utilities)

| Concept ID | Name | Origin |
|------------|------|--------|
| `CONCEPT:ECO-4.0` | Unified Toolkit Ingestion | agent-utilities |
| `CONCEPT:ORCH-1.2` | Confidence-Gated Router | agent-utilities |
| `CONCEPT:OS-5.1` | Prompt Injection Defense | agent-utilities |
| `CONCEPT:OS-5.2` | Cognitive Scheduler | agent-utilities |
| `CONCEPT:OS-5.3` | Guardrail Engine | agent-utilities |
| `CONCEPT:OS-5.4` | Audit Logging | agent-utilities |
| `CONCEPT:KG-2.0` | Knowledge Graph Core | agent-utilities |

## Synergy with agent-utilities

This project integrates with `agent-utilities` via `CONCEPT:ECO-4.0` (Unified Toolkit Ingestion). The `microsoft_agent` MCP server registers its tools with the agent-utilities FastMCP middleware, enabling automatic discovery, telemetry, and Knowledge Graph ingestion of all MSFT-* concepts.
