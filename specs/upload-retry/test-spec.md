# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Verification plan

Governing specification: [spec.md](spec.md).

| Requirement | Verification target |
|---|---|
| `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001` | `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client. |
