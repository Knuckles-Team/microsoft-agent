# Deployment

Microsoft Agent supports an MCP subprocess, an authenticated streamable-HTTP
service, and an optional A2A agent. Deployment configuration supplies all identity,
secret, endpoint, trust, storage, and network values; none are packaged.

## Local MCP subprocess

An MCP client can launch the server over stdio:

```json
{
  "mcpServers": {
    "microsoft-agent": {
      "command": "uvx",
      "args": ["--from", "microsoft-agent[mcp]", "microsoft-mcp"],
      "env": {
        "MCP_TOOL_MODE": "intent",
        "MICROSOFT_TENANT_ID": "<tenant-id>",
        "MICROSOFT_CLIENT_ID": "<client-id>",
        "MICROSOFT_AUTH_MODE": "delegated"
      }
    }
  }
}
```

For a private package index or certificate authority, configure the launcher with
its deployment-owned index and trust profile. Certificate verification remains
enabled.

## Streamable HTTP

Run a long-lived server on a protected interface:

```bash
microsoft-mcp \
  --transport streamable-http \
  --host 127.0.0.1 \
  --port 8000
```

Configure `AUTH_TYPE` and the corresponding Agent Utilities caller-authentication
settings before exposing the listener. Microsoft Graph authentication and MCP
caller authentication are independent boundaries; both must pass.

A remote MCP client should receive only the deployment-owned HTTPS URL and its
caller-authentication profile:

```json
{
  "mcpServers": {
    "microsoft-agent": {
      "url": "https://microsoft-agent.example.invalid/mcp"
    }
  }
}
```

`example.invalid` is a non-resolving documentation name, not a packaged endpoint.

## Containers

The repository's `docker/Dockerfile` contains separate `mcp` and `agent` targets:

```bash
docker build -f docker/Dockerfile --target mcp -t <registry>/microsoft-agent:<version>-mcp .
docker build -f docker/Dockerfile --target agent -t <registry>/microsoft-agent:<version> .
```

Start the MCP image with deployment-owned environment or secret references. Bind
only the required port and place an authenticated TLS reverse proxy or service mesh
in front of any non-loopback listener. Do not put tokens, client secrets,
certificate material, or tenant-specific profiles in an image or Compose file.

The runtime must be non-root, read-only where supported, capability-minimized,
resource-bounded, and covered by an egress allowlist for the configured identity
authority and Graph endpoint.

## A2A agent

The optional A2A process consumes the authenticated MCP service:

```bash
microsoft-agent \
  --mcp-url <mcp-url> \
  --provider <provider> \
  --model-id <model-id> \
  --host 127.0.0.1 \
  --port 9000
```

The provider endpoint, model identifier, credentials, observability settings, TLS
profile, and MCP configuration are external Agent Utilities configuration. Require
the same identity, authorization, retention, and network boundaries used by
GraphOS delegation.

## Identity modes

Select one supported mode with `MICROSOFT_AUTH_MODE`:

| Mode | Required deployment material |
|---|---|
| `delegated` | tenant, client identifier, and approved interactive login method |
| `application` | tenant, client identifier, and certificate or secret reference |
| `on_behalf_of` | confidential client material and verified incoming user token |
| `external_token` | verified request token and explicitly allowed audience |
| `managed_identity` | Azure managed-identity availability |
| `workload_identity` | tenant, client identifier, and federated token-file reference |

Prefer managed identity, workload identity, or a certificate over a static client
secret. Device-code login is disabled until explicitly enabled. Tokens are never
written to a plaintext cache.

## TLS and private certificate authorities

Mount a PEM bundle containing the required intermediate and root certificates and
reference it through the deployment trust profile. Requests-compatible clients can
use `REQUESTS_CA_BUNDLE`; OpenSSL-compatible clients can use `SSL_CERT_FILE`; `uvx`
package resolution can use native platform trust with `UV_NATIVE_TLS=true`.

Do not set an insecure verification mode or package a workstation-specific bundle
path. The operator's Agent Utilities doctor must confirm the selected trust profile
before release validation.

## Optional integrations

Power Platform, Intune, the Windows companion, document roots, Office origins, and
their allowlists are supplied through an external profile referenced by
`MICROSOFT_INTEGRATIONS_CONFIG_PATH`. Checked-in samples define schemas and native
connection points only. They must not contain organization-specific URLs, device
lists, credentials, or customized ontologies.

## Release verification

Before promotion:

1. Validate configuration and identity readiness without printing secret values.
2. Confirm the complete TLS chain with verification enabled.
3. Check that only approved tool groups are registered.
4. Exercise one least-privilege read and one authenticated denial.
5. Validate the connector bundle, schema fingerprints, ontology, mappings, and
   quarantine policy against the installed tool schemas.
6. Confirm metadata-only Langfuse trace linkage and governed graph read-back.
7. Stop the service and verify that no Microsoft Agent or GraphOS process remains.

See [Configuration](configuration.md), [Authentication](authentication.md), and the
[enrollment checklist](enrollment-checklist.md) for the complete operator contract.
