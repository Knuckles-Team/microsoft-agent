# Configuration, trust, and privacy

This page is the operator contract for `microsoft-agent`. Package-specific identity
and policy values are loaded through validated Microsoft settings; shared model,
graph, MCP, trust, observability, and secret references use the Agent Utilities
configuration boundary. Runtime values must be injected by the launcher or an
external AgentConfig profile. They do not belong in source, packaged skill content,
traces, or generated reports.

## Capability configuration

The current capability surface is defined by three versioned artifacts:

- the action-routed MCP tools described in the README and `docs/usage.md`;
- the single canonical `microsoft-agent-operations` skill;
- `connector_manifest.yml` and its ontology, mappings, shapes, fixtures,
  migrations, tool-schema fingerprints, and certification metadata.

Treat those artifacts as a unit during release and deployment. Do not enable a
skill whose certification or tool-schema fingerprint does not match the installed
package. Delegated agents use the current compact/intent-oriented surface; there is
no parallel per-operation compatibility surface.

## AgentConfig and doctor

The deployment profile owns shared configuration, including graph-engine mode and
endpoint, model providers, Langfuse, TLS trust, secret references, workspace policy,
and MCP caller authentication. Microsoft-specific identity and capability keys may
be provided as environment projections of that profile. Doctor output must report
only readiness, selected modes, enabled groups, and opaque configuration origins;
it must never print values, endpoints, credentials, personal paths, or token claims.

Run the Agent Utilities doctor after any configuration change and before starting a
network listener. A missing identity coordinate, unreadable trust reference,
unavailable optional dependency, or invalid integration profile is a failed
readiness gate, not a reason to select a fallback mode.

## Runtime values and secrets

- Supply service endpoints, tenant identifiers, credentials, and model keys
  through environment variables or a mounted secret provider.
- Use non-personal agent aliases and opaque tenant/correlation identifiers.
- Keep developer directories, workstation names, and deployment hostnames out
  of checked-in configuration.
- Bind network transports to an explicitly chosen interface and require the
  deployment's MCP authentication policy before accepting remote traffic.
- Enable optional agent, embedding, evolution, or observability features only
  when their dependencies and backends are configured and healthy.

The checked-in examples use `localhost` for loopback-only development and
`example.invalid` for replaceable network endpoints. Neither value is a
production default.

## TLS trust

Microsoft Graph transport resolves `MICROSOFT_GRAPH_TLS_PROFILE` or
`MICROSOFT_GRAPH_TLS_PROFILE_REF` through the shared Agent Utilities TLS profile
boundary. Power Platform and Windows companion configuration objects accept the
same mutually exclusive `tls_profile` and `tls_profile_ref` fields. Profile names,
secret references, trust material, and service authorities remain deployment
configuration; the package does not contain an environment profile.

Verification and hostname checks remain mandatory for system trust, private
certificate authorities, and mTLS. Provider clients disable redirects and ambient
proxy inheritance and use the shared DNS-pinned egress transport. A private Power
Platform or companion authority is admitted only when it is the exact authority in
the validated external integration configuration. Microsoft Graph and its opaque
upload-session URLs do not receive a checked-in host allowlist; public address
validation, DNS pinning, preserved Host/SNI, and connected-peer verification are
enforced when the connection is opened.

Certificate verification is required. For a private certificate authority,
mount a PEM bundle containing the required intermediate and root certificates,
then configure the client environment with `SSL_CERT_FILE` and, for
Requests-compatible clients, `REQUESTS_CA_BUNDLE`. When `uvx` must use the
native platform trust store while resolving packages, set `UV_NATIVE_TLS=true`.

Do not disable verification to work around an incomplete server chain. Keep CA
bundle locations environment-configured and stable for the runtime; never embed
a workstation path or certificate material in MCP configuration.

## Privacy and data governance

The default observability posture is metadata-only. Do not persist prompts,
message bodies, tool inputs/results, document content, raw traces, credentials,
local paths, hostnames, or personal identity unless an approved data contract
explicitly requires it. Keep Langfuse or OTLP content capture disabled unless a
reviewed retention and access policy authorizes it.

When connector ingestion is enabled, each change must carry tenant, ACL,
classification, retention, provenance, and checkpoint/delta metadata. Reject or
quarantine records that cannot satisfy that contract; never silently widen a
tenant scope. Logs and reports should contain counts, status, and opaque
references only.

Microsoft projection additionally requires
`MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF`, resolved from AgentConfig or a secret
store. It must contain at least 32 bytes and must not be reused as an identity
credential. Only keyed opaque node identifiers and structural relationships may
cross the projection boundary; Microsoft content and identity fields remain
transient.

## Deployment verification

1. Validate the capability bundle and skill metadata against the installed tool
   schemas.
2. Confirm required secrets are present without printing their values.
3. Verify the complete TLS chain with certificate verification enabled.
4. Exercise health/readiness and one least-privilege read operation.
5. Confirm traces arrive under the expected opaque tenant/run identifiers and
   contain no captured content.
6. Record only sanitized pass/fail evidence and version identifiers.
