---
name: microsoft-agent-operations
skill_type: skill
description: >-
  Operate microsoft-agent through its governed MCP and GraphOS capabilities, including discovery, governed operations, ingestion, and verification. Use when a request requires this provider's read, change, automation, ingestion, troubleshooting, or evidence workflows.
---

# Microsoft Agent Operations

Use the provider's governed MCP tools through GraphOS delegation.

## Workflow

1. Establish the verified GraphSession and tenant before discovery or retrieval.
2. Discover the current condensed tool surface; never assume a stale tool name or schema.
3. Prefer read-only inspection first. For changes, present impact and use the provider's
   dry-run or preview mode when available.
4. Execute mutations as fenced WorkItems so retries remain idempotent and auditable.
5. Ingest source data only through the signed connector preset and ChangeEnvelope path.
6. Verify the durable result and its trace/evidence before reporting completion.

## Operation families

- Identity and directory: authentication, users, groups, applications, policy,
  audit, reports, and security.
- Collaboration: mail, calendars, contacts, tasks, Teams, chat, meetings, files,
  sites, Excel, OneNote, and print.
- Optional native connection points: documents, Office bridge, Power Platform,
  Intune, and the outbound Windows companion. Use them only when their dependency,
  external profile, and doctor readiness gate all pass.
- Knowledge projection: `list_microsoft_ingestion_projection` for a read-only keyed
  projection; let GraphOS validate the signed source manifest and commit it
  through the native ChangeEnvelope authority.

The installed MCP schema is authoritative. Route the discovered action through its
current family and preserve the caller's session and approval context.

## Safety contract

- Never persist credentials, endpoints, provider content, raw or unhashed personal
  identifiers, hostnames, filenames, URLs, timestamps, attachment bytes, or local
  paths.
- Microsoft graph projection requires a deployment-owned pseudonymization key of at
  least 32 bytes. The key never appears in tool input, output, traces, or reports.
- Resolve TLS trust and verification from environment/configuration; never hardcode bypasses.
- Treat unknown ACL, tenant, schema, or tool-contract state as a hard failure.
- Require explicit approval for destructive, externally visible, or irreversible actions.
- Keep runtime traces policy-scoped and privacy-sanitized.

## Verification

For every direct or delegated workflow, retain only the exact package/skill/model
identity, pass/fail status, opaque trace reference, governed parent graph reference,
and bounded counts. Verify that the Langfuse record is metadata-only and that the
graph record contains no provider values before claiming success.

## Specialized workflows

Read [the workflow catalog](references/catalog.md) only when the request needs a
provider-specific procedure, parameter map, script, or reference asset.
