# Capability concepts

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

The action catalog covers authentication and metadata; users, groups, directory,
applications, policies, audit, security, and reports; Outlook, calendar, contacts,
tasks, files, sites, Excel, OneNote, Teams, and chat; and supporting device,
storage, print, communications, subscription, and search domains.

The installed MCP schema is the authoritative domain/action inventory. Documentation
describes capability families only so it cannot silently diverge from dynamically
registered tool schemas.

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
