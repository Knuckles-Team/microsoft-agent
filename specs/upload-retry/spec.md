# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Upload retries through the SDK

## Outcome

The connector must retry upload fragments through the shared connector SDK retry policy.
The connector must keep its tool names, signatures, and return values.

## Scope and owner

The `microsoft-agent` repository owns Microsoft Graph and Universal Print upload operations.
The shared connector SDK owns the governed HTTP client and the retry policy.
The connector must not run retry loops of its own.

Component: Connector integration
References: [Centralized capabilities](https://github.com/Knuckles-Team/scion-sdk/blob/main/specs/centralized-capabilities/spec.md), [Transport TLS](../transport-tls/spec.md)
Code: ../../microsoft_agent/power_platform.py ../../microsoft_agent/graph_file_service.py ../../microsoft_agent/universal_print.py ../../microsoft_agent/integration_tools.py ../../tests/test_sdk_centralization.py

## Non-goals

This specification grants no service access and changes no tool contract.
The Entra token verifier in `windows_control_plane.py` keeps its own JWT validation, because it checks tenant-specific claims.

## Why a new spec

The search `git grep -n "retry" -- specs` found no owner for upload retries.
The `transport-tls` spec owns TLS only. This spec owns the retry behavior.

## User Scenarios & Testing

### User Story 1 - SDK upload retries (Priority: P1)

The fleet operator and the engine rely on this behaviour of `microsoft-agent`: The connector must retry upload fragments through the SDK retry policy set in `microsoft_agent/power_platform.py`. It is proven independently of the other stories.

**Acceptance Scenarios**:

1. **Given** the `microsoft-agent` connector built from the default branch, **When** `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` runs, **Then** `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client.

### Edge Cases

- The vendor service is unreachable or rejects the credential: the connector returns the SDK typed error and never a partial result.
- An argument outside the documented inputs is refused with the SDK typed refusal and causes no effect on the vendor service.

## Requirements

- **FR-001**: SDK upload retries. Rows: `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001`.

## Success Criteria

- **SC-001**: `tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client` must observe three attempts through the SDK client. Rows: `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001`.

## Assumptions

- The `microsoft-agent` package keeps the entry point and module layout named in the rows.
- The shared SDK supplies the transport, authentication, retry and certification helpers the rows reuse.

## Traceability

| Label | Rows | Design section | Test ID | Evidence |
|---|---|---|---|---|
| FR-001 | `CONN-MICROSOFT-AGENT-UPLOAD-RETRY-R001` | plan.md Production caller chain | T-001 | tests/test_sdk_centralization.py::test_print_range_retries_through_the_sdk_client |
