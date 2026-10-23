# CONN-MICROSOFT-AGENT-UPLOAD-RETRY — Design and implementation plan

Governing specification: [spec.md](spec.md).

## Existing system and reuse

The connector builds Graph, Power Platform, and companion clients with the SDK HTTP factory.
This plan moves the last hand-written retry loops onto the SDK client.

## Interfaces and live integration path

`microsoft_agent/integration_tools.py` builds the Graph file transport with `upload_retry_policy(settings.max_retries)`.
`submit_print_document` in `microsoft_agent/api/api_client_other.py` builds the print transport with the same policy.
`HttpxAsyncHttpTransport` passes the policy to the SDK client. The SDK waits and honors `Retry-After`.

## Decisions

The SDK retries only idempotent methods. Fragment uploads use PUT, so each fragment retries safely.
The Entra token verifier keeps PyJWT, because the SDK verifiers do not check the companion device claims.
