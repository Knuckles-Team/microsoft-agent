# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Implementation tasks

Governing spec: [spec.md](spec.md). Design: [plan.md](plan.md). Rows: [requirements.md](requirements.md).

- [x] Add `build_client_tls_context` and route both callers through it.
- [x] Prove behavior and the absence of hand-built contexts in `tests/test_transport_tls.py`.
