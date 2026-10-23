# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Implementation tasks

Governing specification: [spec.md](spec.md).

- [x] `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001` — Retry upload fragments through the SDK policy. `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client.
