# Microsoft Agent contributor contract

## Architecture

- `microsoft_agent/mcp_server.py` exposes the current condensed Microsoft Graph
  action tools through the Agent Utilities MCP server factory.
- `microsoft_agent/api_client.py` composes the domain clients in
  `microsoft_agent/api/`; it is the only Graph API client authority.
- `microsoft_agent/auth.py` and `microsoft_agent/settings.py` implement validated
  delegated, application, on-behalf-of, external-token, managed-identity, and
  workload-identity configuration.
- `microsoft_agent/integration_tools.py` supplies optional Office document,
  Power Platform, Intune, and Windows companion tools.
- `microsoft_agent/tool_policy.py` is the fail-closed read/write/destructive
  authorization boundary.
- `connector_manifest.yml`, `microsoft_agent/connectors/`, and
  `microsoft_agent/ontology/` form the provider-owned ingestion and ontology
  capability bundle.
- `microsoft_agent/skills/microsoft-agent-operations/` is the single consolidated
  provider skill.

## Current-only policy

- Do not add compatibility aliases, retired environment variables, duplicate API
  wrappers, duplicate skill families, or fallback execution planes.
- Do not package databases, token caches, MCP client configurations, generated
  runtime state, credentials, endpoints, hostnames, or environment-specific
  profiles.
- Configuration values belong in deployment-owned environment/configuration or
  referenced secret stores. Repository examples must remain generic and value-free.
- TLS verification is mandatory and configured by the deployment trust profile.
  Never hardcode verification bypasses.
- Use the Agent Utilities server/session/configuration boundaries; do not create a
  second MCP, authentication, graph, or agent runtime.

## Security invariants

- Microsoft identity coordinates are mandatory at authentication time; there is no
  baked-in client identifier.
- Tokens are stored only through secure OS-backed storage. A keyring failure leaves
  tokens in memory and must not enable plaintext persistence.
- Condensed action tools authorize the routed action, not merely the envelope tool
  name. Unknown actions are treated as writes and fail closed.
- Writes require `MICROSOFT_ALLOW_WRITES`; destructive actions additionally require
  `MICROSOFT_ALLOW_DESTRUCTIVE`.
- Connector materialization remains quarantined until tenant, ACL, provenance,
  schema, and privacy policy are verified.
- Tests must mock all provider and network interactions.

## Development commands

```bash
uv sync --all-extras --dev
uv run pytest
uv run ruff check .
uv run python scripts/security_sanitizer.py
uv run python scripts/security_contract.py --contract .security/security-contract.json validate
uv run python scripts/verify_api_integration.py --local
uv run mkdocs build --strict
```

Run commands from the repository root. Do not run native builds or services in
parallel with other resource-intensive workspace validation.

## Change discipline

- Preserve unrelated user changes and repository history.
- Use a dedicated branch or worktree per concurrent session; configure workspace and
  worktree roots outside the repository.
- Do not reset, stash, rebase, force-push, or discard another session's work.
- Keep versions, dependency constraints, lockfiles, manifests, generated provider
  attestations, documentation, and tests synchronized.
- A change is complete only when its focused tests, static checks, security/privacy
  gates, documentation contract, and generated artifacts agree.
