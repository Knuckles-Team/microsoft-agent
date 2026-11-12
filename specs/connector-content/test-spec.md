# CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT — Test specification

Governing spec: [spec.md](spec.md).

| Test ID | Requirement | Proving test | Expected observation |
|---|---|---|---|
| T-001 | `CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT-R001` | `tests/test_default_content_surfaces.py::test_default_content_surfaces` | `resources/list` includes `skill://` and `ontology://`; `prompts/list` includes each packaged prompt. |
| T-002 | `CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT-R002` | `tests/test_destructive_declaration.py::test_destructive_tools_are_declared_and_take_confirm` | Each destructive tool is served with `destructiveHint` true and a `confirm` argument, none has an allow argument, and a hand-written destructive tool refuses without confirmation and sends no request. |

## Negative and boundary cases

Each proving test above includes the refusal or error case that its row names. A vendor outage or an unauthorised call must return the SDK typed error.

## Quality and release proof

Run `uv run pytest tests` and `pre-commit run --all-files` from the repository root. A planned test is added by the code change that lands the row.
