# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Implementation plan

Governing spec: [spec.md](spec.md). Rows: [requirements.md](requirements.md).

## Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-UPLOAD-RETRY`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

## Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 1 active row(s) in this spec

## Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

## Project Structure

- `microsoft_agent/power_platform.py` (existing)
- `microsoft_agent/graph_file_service.py` (existing)
- `microsoft_agent/universal_print.py` (existing)
- `microsoft_agent/mcp_server.py` (existing)
- `tests/test_sdk_centralization.py` (existing)

## Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

## Phase 1: Design

### Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-UPLOAD-RETRY`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

### Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 1 active row(s) in this spec

### Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

### Project Structure

- `microsoft_agent/power_platform.py` (existing)
- `microsoft_agent/graph_file_service.py` (existing)
- `microsoft_agent/universal_print.py` (existing)
- `microsoft_agent/mcp_server.py` (existing)
- `tests/test_sdk_centralization.py` (existing)

### Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

### Phase 1: Design

### Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-UPLOAD-RETRY`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

### Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 1 active row(s) in this spec

### Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

### Project Structure

- `microsoft_agent/power_platform.py` (existing)
- `microsoft_agent/graph_file_service.py` (existing)
- `microsoft_agent/universal_print.py` (existing)
- `microsoft_agent/mcp_server.py` (existing)
- `tests/test_sdk_centralization.py` (existing)

### Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

### Phase 1: Design

### Summary

This plan states how `microsoft-agent` delivers the rows of `CONN-MICROSOFT-AGENT-UPLOAD-RETRY`. The live caller is the `microsoft-mcp` console script. The owner is the `microsoft_agent` package. The rows reuse the `scion-sdk` helpers and add no second authority. The plan changes no runtime code.

### Technical Context

- **Language/Version**: Python >=3.12
- **Primary Dependencies**: scion-sdk[mcp]>=0.6.4,<0.7.0; fastmcp; further pins in `pyproject.toml` (scion-sdk[mcp]>=0.6.4,<0.7.0, azure-core>=1.41.0, httpx>=0.28.1, keyring>=25.7.0, msal>=1.39.0, msgraph-sdk>=1.64.0)
- **Storage**: N/A, the connector keeps no store of its own
- **Testing**: pytest; each row names the proving test as `path::name`
- **Target Platform**: Linux container and stdio or streamable HTTP MCP clients
- **Project Type**: connector
- **Performance Goals**: N/A, the vendor service sets the latency
- **Constraints**: vendor calls go through the SDK governed client; no connector-local transport
- **Scale/Scope**: 1 active row(s) in this spec

### Constitution Check

- Organization constitution 1.5.0, spec-first change: pass. Every behaviour is a row of this spec with a bound or planned test.
- Repository article R-I, connector scope: pass. The rows stay inside this connector's one external system and reuse the SDK for shared capabilities.

### Project Structure

- `microsoft_agent/power_platform.py` (existing)
- `microsoft_agent/graph_file_service.py` (existing)
- `microsoft_agent/universal_print.py` (existing)
- `tests/test_sdk_centralization.py` (existing)

### Phase 0: Research

No unknowns. Every row names its code and test; the caller chain of each row was read from the code.

### Phase 1: Design

Governing specification: [spec.md](spec.md).

### Existing system and reuse

The connector builds Graph, Power Platform, and companion clients with the SDK HTTP factory.
This plan moves the last hand-written retry loops onto the SDK client.

### Interfaces and live integration path

`microsoft_agent/integration_tools.py` builds the Graph file transport with `upload_retry_policy(settings.max_retries)`.
`submit_print_document` in `microsoft_agent/api/api_client_other.py` builds the print transport with the same policy.
`HttpxAsyncHttpTransport` passes the policy to the SDK client. The SDK waits and honors `Retry-After`.

### Decisions

The SDK retries only idempotent methods. Fragment uploads use PUT, so each fragment retries safely.
The Entra token verifier keeps PyJWT, because the SDK verifiers do not check the companion device claims.

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

- `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001`: an upload tool of `microsoft_agent/mcp_server.py` calls `graph_file_service.py` or `universal_print.py`, which send each fragment through the SDK client built with `upload_retry_policy` (`microsoft_agent/power_platform.py:119`).

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

- `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001`: an upload tool of `microsoft_agent/mcp_server.py` calls `graph_file_service.py` or `universal_print.py`, which send each fragment through the SDK client built with `upload_retry_policy` (`microsoft_agent/power_platform.py:119`).

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

- `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001`: an upload tool of `microsoft_agent/mcp_server.py` calls `graph_file_service.py` or `universal_print.py`, which send each fragment through the SDK client built with `upload_retry_policy` (`microsoft_agent/power_platform.py:119`).

## Cross-repo seams

None. The rows depend only on the published `scion-sdk` package, which the connector imports.
