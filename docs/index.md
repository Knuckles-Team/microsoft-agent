# Microsoft Agent

Microsoft Agent is a governed Microsoft Graph MCP provider and optional A2A agent
for the Agent Utilities and Epistemic Graph ecosystem.

It supplies:

- one modular asynchronous Graph client;
- a compact intent-oriented MCP surface backed by action-routed provider tools;
- delegated, application, on-behalf-of, external-token, managed-identity, and
  workload-identity authentication;
- fail-closed read, write, and destructive action policy;
- optional document, Office bridge, Power Platform, Intune, and Windows companion
  connection points;
- governed Epistemic Graph ingestion with provider-owned ontology, mappings,
  provenance, schema fingerprints, and quarantine; and
- one consolidated `microsoft-agent-operations` skill.

The package contains no tenant profile, credentials, token cache, database, local
filesystem configuration, endpoint inventory, or customized ontology.

## Start here

- [Architecture](overview.md) explains runtime authorities and data flow.
- [Installation](installation.md) lists package extras and verification.
- [Configuration](configuration.md) defines AgentConfig, trust, and privacy rules.
- [Authentication](authentication.md) covers identity modes and permission profiles.
- [Usage](usage.md) documents the current MCP, Python, and A2A interfaces.
- [Deployment](deployment.md) gives environment-neutral process and container
  patterns.
- [Integrations](integrations.md) describes optional native connection points.

## Minimal stdio launch

```bash
python -m pip install "microsoft-agent[mcp]"
MICROSOFT_TENANT_ID=<tenant-id> \
MICROSOFT_CLIENT_ID=<client-id> \
microsoft-mcp --transport stdio
```

All runtime identity, trust, model, graph, observability, and optional integration
values are deployment-owned. See [Configuration](configuration.md) before enabling
writes or a network listener.
