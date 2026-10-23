# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Design and implementation plan

Governing spec: [spec.md](spec.md).

## Existing system and reuse

Both `windows_companion_service.py` and `windows_control_plane.py` built a context with the standard library. The SDK `resolve_tls_profile` already owns trust, minimum version and proxy settings.

## Architecture

`build_client_tls_context` in `windows_runtime.py` calls `resolve_tls_profile("microsoft", profile=...)`. With a CA bundle it passes `{"ca_bundle_path": ...}`; without one the deployment's `MICROSOFT_*` profile settings apply. The key password is a runtime value a profile document may not hold, so the certificate chain is loaded onto the returned context.

## Interfaces and data model

No wire, schema or configuration change.

## Live integration path

`build_worker` and `HttpOutboundRelayTransport` call the helper; the start-up validation calls it to prove the certificate and key load.

## Quality and release gates

`uv run pytest tests/test_transport_tls.py tests/test_windows_control_plane.py tests/test_windows_runtime.py`, `uv run ruff check .`, pre-commit.

## Risks, alternatives, and decisions

A profile with an encrypted key would need a secret reference; the runtime password cannot be inlined, hence the post-resolution load.
