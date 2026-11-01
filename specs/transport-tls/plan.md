# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Implementation plan

Governing spec: [spec.md](spec.md). Rows: [requirements.md](requirements.md).

## Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-TRANSPORT-TLS`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

## Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 2 active row(s) in this spec

## Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

## Project Structure

- `microsoft_agent/windows_runtime.py` (existing)
- `microsoft_agent/windows_companion_service.py` (existing)
- `microsoft_agent/windows_control_plane.py` (existing)
- `tests/test_transport_tls.py` (existing)

## Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

## Phase 1: Design

### Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-TRANSPORT-TLS`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

### Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 2 active row(s) in this spec

### Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

### Project Structure

- `microsoft_agent/windows_runtime.py` (existing)
- `microsoft_agent/windows_companion_service.py` (existing)
- `microsoft_agent/windows_control_plane.py` (existing)
- `tests/test_transport_tls.py` (existing)

### Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

### Phase 1: Design

### Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-TRANSPORT-TLS`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

### Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 2 active row(s) in this spec

### Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

### Project Structure

- `microsoft_agent/windows_runtime.py` (existing)
- `tests/test_transport_tls.py` (existing)

### Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

### Phase 1: Design

### Existing system and reuse

Both `windows_companion_service.py` and `windows_control_plane.py` built a context with the standard library. The SDK `resolve_tls_profile` already owns trust, minimum version and proxy settings.

### Architecture

`build_client_tls_context` in `windows_runtime.py` calls `resolve_tls_profile("microsoft", profile=...)`. With a CA bundle it passes `{"ca_bundle_path": ...}`; without one the deployment's `MICROSOFT_*` profile settings apply. The key password is a runtime value a profile document may not hold, so the certificate chain is loaded onto the returned context.

### Interfaces and data model

No wire, schema or configuration change.

### Live integration path

`build_worker` and `HttpOutboundRelayTransport` call the helper; the start-up validation calls it to prove the certificate and key load.

### Quality and release gates

`uv run pytest tests/test_transport_tls.py tests/test_windows_control_plane.py tests/test_windows_runtime.py`, `uv run ruff check .`, pre-commit.

### Risks, alternatives, and decisions

A profile with an encrypted key would need a secret reference; the runtime password cannot be inlined, hence the post-resolution load.

### Phase 2: Task approach

`tasks.md` holds one task per row, producer first. Rows that depend on none run in parallel. A task is checked when its `Spec:` trailer is on the default branch.

### Complexity Tracking

None. The constitution check passed.

### Build contract

- `pre-commit run --all-files` runs ruff check, ruff format, the shared `connector-standard`, `spec`, `secrets` and `dependency-readiness` gates at the pinned shared revision.
- `uv run pytest tests` runs the proving tests; the shared complexity, duplication and KISS gates apply their org thresholds and this repository sets no local override.
- A change adds no scanner suppression and no duplicate implementation of an SDK helper.

### Production caller chain


### Cross-repo seams

None. The rows depend only on the published `scion-sdk` package, which the connector imports.

### Phase 2: Task approach

`tasks.md` holds one task per row, producer first. Rows that depend on none run in parallel. A task is checked when its `Spec:` trailer is on the default branch.

### Complexity Tracking

None. The constitution check passed.

### Build contract

- `pre-commit run --all-files` runs ruff check, ruff format, the shared `connector-standard`, `spec`, `secrets` and `dependency-readiness` gates at the pinned shared revision.
- `uv run pytest tests` runs the proving tests; the shared complexity, duplication and KISS gates apply their org thresholds and this repository sets no local override.
- A change adds no scanner suppression and no duplicate implementation of an SDK helper.

### Production caller chain

- `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001`: `windows_companion_service.py` (`microsoft_agent/windows_companion_service.py:252`) and `windows_control_plane.py` (`microsoft_agent/windows_control_plane.py:1238`) call `build_client_tls_context` (`microsoft_agent/windows_runtime.py:1305`), which resolves the SDK TLS profile.
- `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002`: the package scan test walks every module under `microsoft_agent/` and fails on `ssl.create_default_context`, `ssl.SSLContext` or `verify=False`.

### Cross-repo seams

None. The rows depend only on the published `scion-sdk` package, which the connector imports.

## Phase 2: Task approach

`tasks.md` holds one task per row, producer first. Rows that depend on none run in parallel. A task is checked when its `Spec:` trailer is on the default branch.

## Complexity Tracking

None. The constitution check passed.

## Build contract

- `pre-commit run --all-files` runs ruff check, ruff format, the shared `connector-standard`, `spec`, `secrets` and `dependency-readiness` gates at the pinned shared revision.
- `uv run pytest tests` runs the proving tests; the shared complexity, duplication and KISS gates apply their org thresholds and this repository sets no local override.
- A change adds no scanner suppression and no duplicate implementation of an SDK helper.

## Production caller chain

- `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001`: `windows_companion_service.py` (`microsoft_agent/windows_companion_service.py:252`) and `windows_control_plane.py` (`microsoft_agent/windows_control_plane.py:1238`) call `build_client_tls_context` (`microsoft_agent/windows_runtime.py:1305`), which resolves the SDK TLS profile.
- `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002`: the package scan test walks every module under `microsoft_agent/` and fails on `ssl.create_default_context`, `ssl.SSLContext` or `verify=False`.

## Cross-repo seams

None. The rows depend only on the published `scion-sdk` package, which the connector imports.
