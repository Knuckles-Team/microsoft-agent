# Architecture overview

Microsoft Agent is a governed Microsoft Graph provider for the Agent Utilities and
Epistemic Graph ecosystem. It contributes one MCP server, one modular Graph client,
one authentication authority, optional integration adapters, a consolidated skill,
and a provider-owned connector capability bundle.

## Runtime boundaries

| Boundary | Authority |
|---|---|
| MCP server and session | Agent Utilities server factory and verified caller session |
| Microsoft Graph client | `MicrosoftGraphApi` composed from `microsoft_agent/api/` |
| Microsoft authentication | `microsoft_agent/auth.py` and validated `MicrosoftSettings` |
| Action authorization | `MicrosoftToolPolicy` applied to the routed action |
| Agent delegation | Agent Utilities A2A/GraphOS runtime |
| Knowledge ingestion | governed connector manifest, ontology, mapping, and quarantine |
| Observability | deployment-configured metadata-only Agent Utilities telemetry |

There is no second API wrapper, raw-token fallback, plaintext token cache, bundled
database, bundled MCP profile, or alternate agent execution plane.

## Capability flow

1. The deployment supplies Microsoft identity, MCP caller identity, TLS trust, and
   an explicit optional-capability profile.
2. The MCP server creates a verified Agent Utilities session and exposes the
   configured action families through the selected tool mode.
3. The tool policy classifies the routed action. Reads are allowed by default;
   writes and destructive actions require separate explicit approvals.
4. The single Graph client acquires a token through the selected Microsoft identity
   mode and performs the bounded provider request.
5. Governed ingestion converts approved records into quarantined change envelopes
   carrying tenant, ACL, provenance, schema, checkpoint, and privacy metadata.
6. Metadata-only telemetry links the MCP/delegation outcome to its Langfuse trace
   and governed parent graph record.

## Source layout

```text
microsoft_agent/
  api/                    domain Graph client mixins
  connectors/             provider source presets and schema fingerprints
  ontology/               ontology, shapes, mappings, fixtures, and migrations
  skills/                 consolidated Microsoft operations skill
  api_client.py           sole Graph API client
  auth.py                 token acquisition and secure-cache authority
  settings.py             validated identity and policy configuration
  mcp_server.py           condensed/action-routed MCP server
  tool_policy.py          fail-closed action policy
  integration_tools.py    optional native connection points
  agent_server.py         optional Agent Utilities A2A entry point
```

Deployment and Office add-in assets define generic schemas and packaging only.
Environment URLs, credentials, device inventories, filesystem roots, trust bundles,
and customized ontologies stay outside the repository.

## Security model

- Microsoft Graph identity and MCP caller identity are independently verified.
- Unknown routed actions fail closed as writes.
- Side-effecting and destructive actions use separate enablement controls.
- Secure OS-backed storage is the only persistent delegated-token cache.
- Certificate verification remains enabled and trust is deployment-configured.
- Provider results are bounded, policy-checked, and privacy-sanitized before
  persistence or observability.
- Connector records remain quarantined until their governance contract passes.

## Extension model

New provider operations belong in the relevant domain client and action family.
New optional systems belong behind a native adapter registered only when its
package extra and external configuration are both present. New ingestion semantics
extend the provider ontology and mapping, followed by regeneration of schema
fingerprints and certification artifacts.

Do not add per-operation compatibility tools, alternate authentication managers,
hardcoded endpoints, local profiles, or provider-specific execution runtimes.
