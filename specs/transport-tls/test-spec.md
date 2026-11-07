# CONN-MICROSOFT-AGENT-TRANSPORT-TLS — Test specification

Governing spec: [spec.md](spec.md).

| Test ID | Requirement | Proving test | Expected observation |
|---|---|---|---|
| T-001 | `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001` | `microsoft_agent/windows_runtime.py::build_client_tls_context` | The context verifies peer and hostname and trusts the instance CA. |
| T-002 | `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002` | `tests/test_transport_tls.py::test_no_hand_built_tls_context_in_package` | The scan finds no offender. |

## Negative and boundary cases

Each proving test above includes the refusal or error case that its row names. A vendor outage or an unauthorised call must return the SDK typed error.

## Quality and release proof

Run `uv run pytest tests` and `pre-commit run --all-files` from the repository root. A planned test is added by the code change that lands the row.

## Adoption of scion-sdk 0.10.0 (`CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R003`)

| Requirement | Test | Observation |
|---|---|---|
| `CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R003` | `tests/test_scion_sdk_0_10_imports.py::test_every_scion_sdk_import_resolves_on_the_pinned_release` | Every `scion_sdk` import resolves on the pinned release; a removed path fails the test. |
