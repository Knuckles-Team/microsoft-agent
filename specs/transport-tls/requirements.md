# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Requirements

Governing spec: [spec.md](spec.md).

Each row is one requirement. A row names the file it touches and the test that proves it. The generator sets each row's state in `status.json`.

| ID | Requirement | Acceptance |
|---|---|---|
| `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001` | **Profile-built context.** `microsoft_agent/windows_runtime.py::build_client_tls_context` resolves trust through the SDK `resolve_tls_profile`; a per-instance CA goes through the profile and the client certificate loads onto its context. A missing bundle raises. Caller chain: `windows_companion_service.py` (`microsoft_agent/windows_companion_service.py:252`) and `windows_control_plane.py` (`microsoft_agent/windows_control_plane.py:1238`) call `build_client_tls_context` (`microsoft_agent/windows_runtime.py:1305`), which resolves the SDK TLS profile. Shape: `build_client_tls_context` in `microsoft_agent/windows_runtime.py`; reuses `resolve_tls_profile` of the SDK. Proving test: `tests/test_transport_tls.py::test_client_context_verifies_peer_and_uses_instance_ca`. | The context verifies peer and hostname and trusts the instance CA. |
| `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002` | **No hand-built contexts.** No module under `microsoft_agent/` calls `ssl.create_default_context` or `ssl.SSLContext`, or passes `verify=False`. Caller chain: the package scan test walks every module under `microsoft_agent/` and fails on `ssl.create_default_context`, `ssl.SSLContext` or `verify=False`. Shape: the scan in `tests/test_transport_tls.py`; no production code change. Proving test: `tests/test_transport_tls.py::test_no_hand_built_tls_context_in_package`. | The scan finds no offender. |
