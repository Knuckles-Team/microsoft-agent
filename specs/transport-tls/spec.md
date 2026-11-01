# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Outbound transport TLS

Owner: microsoft-agent. Cross-repo IDs: SCION-CONNECTOR-CONTROL-R017, SCION-CENTRALIZED-CAPABILITIES-R008.
Public provenance: the SDK centralization audit (HTTP client, retry, TLS).

## Purpose and user stories

An operator runs the Windows companion and its control-plane relay. Every outbound TLS connection must be verified by the one trust profile the deployment configures, with any per-instance CA bundle and client certificate carried through that profile.

## Why a new spec

The repository had no local spec. `specs/README.md` and `specs/_template` were read; no spec owns companion or relay transport.

## Scope and interfaces

Owned: `microsoft_agent/windows_runtime.py` (`build_client_tls_context`), its two callers `windows_companion_service.py` and `windows_control_plane.py`. The profile format and verification invariants belong to scion-sdk. Non-goal: HTTP client replacement.

## Open questions

- None.

## Delivery mapping

Component: microsoft-agent connector

Code: ../../microsoft_agent/__init__.py

## User Scenarios & Testing

### User Story 1 - Profile-built context (Priority: P1)

The fleet operator and the engine rely on this behaviour of `microsoft-agent`: `microsoft_agent/windows_runtime.py::build_client_tls_context` resolves trust through the SDK `resolve_tls_profile`; a per-instance CA goes through the profile and the client certificate loads onto its context. It is proven independently of the other stories.

**Acceptance Scenarios**:

1. **Given** the `microsoft-agent` connector built from the default branch, **When** `microsoft_agent/windows_runtime.py::build_client_tls_context` runs, **Then** the context verifies peer and hostname and trusts the instance CA.

### User Story 2 - No hand-built contexts (Priority: P1)

The fleet operator and the engine rely on this behaviour of `microsoft-agent`: No module under `microsoft_agent/` calls `ssl.create_default_context` or `ssl.SSLContext`, or passes `verify=False`. It is proven independently of the other stories.

**Acceptance Scenarios**:

1. **Given** the `microsoft-agent` connector built from the default branch, **When** `tests/test_transport_tls.py::test_no_hand_built_tls_context_in_package` runs, **Then** the scan finds no offender.

### Edge Cases

- The vendor service is unreachable or rejects the credential: the connector returns the SDK typed error and never a partial result.
- An argument outside the documented inputs is refused with the SDK typed refusal and causes no effect on the vendor service.

## Requirements

- **FR-001**: Profile-built context. Rows: `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001`.
- **FR-002**: No hand-built contexts. Rows: `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002`.

## Success Criteria

- **SC-001**: The context verifies peer and hostname and trusts the instance CA. Rows: `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001`.
- **SC-002**: The scan finds no offender. Rows: `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002`.

## Assumptions

- The `microsoft-agent` package keeps the entry point and module layout named in the rows.
- The shared SDK supplies the transport, authentication, retry and certification helpers the rows reuse.

## Traceability

| Label | Rows | Design section | Test ID | Evidence |
|---|---|---|---|---|
| FR-001 | `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001` | plan.md Production caller chain | T-001 | microsoft_agent/windows_runtime.py::build_client_tls_context |
| FR-002 | `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002` | plan.md Production caller chain | T-002 | tests/test_transport_tls.py::test_no_hand_built_tls_context_in_package |
