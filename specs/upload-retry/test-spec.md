# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Test specification

Governing spec: [spec.md](spec.md).

| Test ID | Requirement | Proving test | Expected observation |
|---|---|---|---|
| T-001 | `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001` | `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` | `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client. |

## Negative and boundary cases

Each proving test above includes the refusal or error case that its row names. A vendor outage or an unauthorised call must return the SDK typed error.

## Quality and release proof

Run `uv run pytest tests` and `pre-commit run --all-files` from the repository root. A planned test is added by the code change that lands the row.
