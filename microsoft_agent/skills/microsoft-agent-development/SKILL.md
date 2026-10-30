---
name: microsoft-agent-development
skill_type: skill
domain: development
description: >-
  Develop, test, gate, and release the microsoft-agent connector. Use this skill
  when changing Microsoft Graph integrations, governed MCP tools, packaged
  connector content, documentation, or repository gates.
license: MIT
tags: [microsoft, graph, development, connector, test, gates]
tier: repository
metadata:
  author: Knuckles-Team
  version: '0.1.0'
---

# microsoft-agent development

This repository owns the Microsoft Graph client, authentication boundary,
action-routed MCP tools, provider mappings, ontology, and packaged connector
content. Shared connector contracts and transports belong to
`scion-sdk`; admission, serving, and durable graph state belong to
their owning repositories. Read `AGENTS.md` before changing code. Keep tenant
identifiers, credentials, private endpoints, provider records, and raw
personal data out of source, logs, fixtures, and packaged skill content.

## Rapid delivery: the 5-minute contract

Specs are already designed. Deliver small, reviewable slices and cite the
owning spec row in the commit trailer.

- Find the owning row first with `git grep -n "<term>" -- specs` and extend the
  existing spec before adding requirement text.
- Reuse the one SDK `Connector`, the existing Graph client, auth/settings
  boundary, and action-routed tool modules. Do not create a second MCP,
  authentication, graph-ingest, or safety plane.
- Preserve fail-closed read/write/destructive policy, tenant and provenance
  checks, TLS verification, and the existing condensed and verbose tool
  surfaces. Provider calls remain mocked in local tests.
- Bind every delivered requirement to a focused test. Generated
  `specs/*/status.json` files come from the shared hooks; never hand-edit them.
- Keep `skills/**`, `ontology/**`, `connectors/**`, prompts, manifests, and
  package-data declarations coherent. Put `Spec: <ID>` and the assisting-agent
  trailer on consecutive final commit message lines.

## Setup and gates

Use the pinned environment and shared gates from the checkout root:

```bash
uv sync --locked
uvx pre-commit run --all-files
uvx pre-commit run --all-files --hook-stage pre-push
```

For focused feedback, use tests that exercise the actual connector paths:

```bash
/var/tmp/arbora-delivery/scale/run-narrow-test-slot.sh \
  uv run --locked --extra test pytest -q \
    tests/test_publication_workflow.py \
    tests/test_default_content_surfaces.py \
    tests/test_startup.py \
    tests/test_mcp_server.py \
    tests/test_auth.py
```

Do not call Microsoft Graph or require tenant credentials in local gates.
Preserve the existing Graph client composition, MSAL identity modes, optional
Office/Power Platform/Intune/Windows integrations, and governed KG ingestion
when changing a provider module.

## Packaging and publication

Keep `microsoft_agent/skills/`, `microsoft_agent/ontology/`,
`microsoft_agent/connectors/`, prompts, and manifests in the package-data and
entry-point contracts. Build locally with `uv build` and inspect the wheel for
`microsoft_agent/skills/microsoft-agent-development/SKILL.md` before a release.
The canonical Pages publisher is `.github/workflows/pages.yml`; it calls the
shared Pipelines workflow with `.config/spec-dashboard.json`, keeps manual
dispatch and the docs/spec/config/workflow path filters, and publishes the
source-backed dashboard at `spec-delivery/`. Do not add a second documentation
publisher or claim Pages publication from a local build.

Keep package version mirrors and `uv.lock` coherent for a genuine packaged
content release. Central generation owns connector manifests and release
container artifacts; preserve this package's existing runtime, auth, and port
semantics while inspecting generated output before making claims.

## Change discipline

- Keep Microsoft Graph calls, mappings, ontology, source declarations, and MCP
  schemas under this repository's connector boundary. SDK transport, shared
  ingest contracts, and Arbora admission belong to their owner repositories.
- Use deployment-owned `MICROSOFT_*` identity and policy configuration only;
  never print or commit credentials, token caches, private endpoints, or tenant
  profiles. Preserve mandatory TLS verification and fail-closed authorization.
- For mutations, inspect current state first, require explicit approval where
  the existing procedure calls for it, and verify the result. Keep KG
  ingestion governed by the signed source and ChangeEnvelope path.
- Keep generated status changes and central promotion with the Pipelines/status
  owner. Do not claim a row is landed or verified from a local edit alone.
