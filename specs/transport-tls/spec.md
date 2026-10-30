# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Outbound transport TLS

Owner: microsoft-agent. Cross-repo IDs: SCION-CONNECTOR-CONTROL-R017, SCION-CENTRALIZED-CAPABILITIES-R008.
Public provenance: the SDK centralization audit (HTTP client, retry, TLS).

## Purpose and user stories

An operator runs the Windows companion and its control-plane relay. Every outbound TLS connection must be verified by the one trust profile the deployment configures, with any per-instance CA bundle and client certificate carried through that profile.

## Requirements and acceptance

- FR-001: The control-plane relay client and the companion start-up validation obtain their TLS context from the SDK `resolve_tls_profile`, never from a hand-built context.
- FR-002: A per-instance CA bundle is passed through the profile; the client certificate and key are loaded onto the profile's context. A missing or invalid bundle fails closed.
- SC-001: No package module builds its own context or disables verification.

## Why a new spec

The repository had no local spec. `specs/README.md` and `specs/_template` were read; no spec owns companion or relay transport.

## Scope and interfaces

Owned: `microsoft_agent/windows_runtime.py` (`build_client_tls_context`), its two callers `windows_companion_service.py` and `windows_control_plane.py`. The profile format and verification invariants belong to scion-sdk. Non-goal: HTTP client replacement.

## Traceability

| Requirement | Design section | Test ID | Evidence |
|---|---|---|---|
| FR-001 | plan.md | tests/test_transport_tls.py | |
| FR-002 | plan.md | tests/test_transport_tls.py | |
| SC-001 | plan.md | tests/test_transport_tls.py | |

## Open questions

- None.
