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

## Requirements

See [requirements.md](requirements.md), [plan.md](plan.md), and [test-spec.md](test-spec.md).
