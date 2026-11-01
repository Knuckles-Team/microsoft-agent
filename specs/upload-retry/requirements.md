# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Requirements

Governing specification: [spec.md](spec.md).

Each row identifies one connector behavior and its verification target.

| ID | Requirement | Acceptance |
|---|---|---|
| `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001` | **SDK upload retries.** The connector must retry upload fragments through the SDK retry policy set in `microsoft_agent/power_platform.py`. `microsoft_agent/graph_file_service.py` and `microsoft_agent/universal_print.py` must send each fragment once. Caller chain: an upload tool of `microsoft_agent/mcp_server.py` calls `graph_file_service.py` or `universal_print.py`, which send each fragment through the SDK client built with `upload_retry_policy` (`microsoft_agent/power_platform.py:119`). Shape: `upload_retry_policy` in `microsoft_agent/power_platform.py` returning an SDK `RetryPolicy`; the two services hold no retry loop. Proving test: `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client`. | `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client. |
