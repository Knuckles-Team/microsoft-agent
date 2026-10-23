# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Requirements

Governing spec: [spec.md](spec.md).

Each row is one requirement. A row names the file it touches and the test that proves it. The generator sets each row's state in `status.json`.

| ID | Requirement | Acceptance |
|---|---|---|
| `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001` | **Profile-built context.** `microsoft_agent/windows_runtime.py::build_client_tls_context` resolves trust through the SDK `resolve_tls_profile`; a per-instance CA goes through the profile and the client certificate loads onto its context. A missing bundle raises. Proving tests: `tests/test_transport_tls.py::test_client_context_verifies_peer_and_uses_instance_ca`, `::test_client_context_uses_the_sdk_profile`, `::test_missing_ca_bundle_fails_closed`. | The context verifies peer and hostname and trusts the instance CA. |
| `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002` | **No hand-built contexts.** No module under `microsoft_agent/` calls `ssl.create_default_context` or `ssl.SSLContext`, or passes `verify=False`. Proving test: `tests/test_transport_tls.py::test_no_hand_built_tls_context_in_package`. | The scan finds no offender. |
