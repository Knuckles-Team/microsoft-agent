# Provider workflow catalog

Load only the section relevant to the current request. Discover exact tool schemas
from the installed MCP server before using any family.

## Read workflow

1. Confirm verified caller and Microsoft identity readiness.
2. Select the smallest matching action family and bounded field set.
3. Execute the read without enabling write policy.
4. Return only the requested data; keep traces metadata-only.

## Write workflow

1. Read the current target and present the intended impact.
2. Require the write policy and caller approval.
3. Supply an idempotency/correlation key when the action supports it.
4. Execute once, then verify with an independent bounded read.

## Destructive workflow

1. Complete the write workflow preflight.
2. Require both write and destructive policy and explicit confirmation bound to the
   exact target and operation.
3. Execute as a fenced WorkItem where supported and reject late/replayed completion.
4. Verify deletion or terminal state without returning removed content.

## Zero-PII graph workflow

1. Confirm the pseudonymization-key and graph-session doctor gates.
2. Use `list_microsoft_ingestion_projection` to inspect only opaque nodes and
   structural relationships.
3. Confirm no provider identifier or content appears in the result.
4. Delegate manifest validation and ChangeEnvelope materialization to GraphOS.
5. Read back the governed parent record and verify tenant, ACL, provenance,
   quarantine, signature, and privacy metadata without revealing values.
