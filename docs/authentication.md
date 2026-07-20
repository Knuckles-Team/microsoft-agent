# Authentication and permissions

Authentication is intentionally configured after installation. The server can
start, expose health/configuration tools, generate local Office documents, and
report integration readiness without a tenant ID or credential.

## Choose one identity mode

| Mode | Intended use | Credential material |
|---|---|---|
| `delegated` | A person using Outlook, Teams, Word, PowerPoint, OneDrive, and SharePoint from a Windows laptop | Windows broker or browser session; secure OS token cache |
| `application` | A background service using tenant-wide or explicitly targeted resources | Certificate preferred; client secret supported |
| `on_behalf_of` | An authenticated API calling Graph for its signed-in user | Incoming user assertion plus confidential app credential |
| `managed_identity` | Azure-hosted service | System- or user-assigned managed identity; no stored secret |
| `workload_identity` | Federated Kubernetes/CI workload | Entra federated token file; no stored secret |
| `external_token` | A trusted front end already supplies a Graph token | Incoming JWT must match the configured tenant and Graph audience |

Graph endpoints under `/me` require a delegated user context (delegated,
on-behalf-of, or an equivalent delegated external token). Application, managed,
and workload identities should use tenant-level tools or explicit-resource
variants such as the shared-mailbox `/users/{id}` tools; the agent must not
pretend that an app-only token has a current user.

For a Windows user, start with delegated broker authentication:

```powershell
$env:MICROSOFT_TENANT_ID = '<tenant-guid>'
$env:MICROSOFT_CLIENT_ID = '<application-client-id>'
$env:MICROSOFT_AUTH_MODE = 'delegated'
$env:MICROSOFT_LOGIN_METHOD = 'auto'
$env:MICROSOFT_ENABLE_BROKER = 'true'
microsoft-mcp --transport stdio
```

`auto` prefers the Windows Web Account Manager broker on native Windows and the
system browser elsewhere. Device-code login is disabled unless both
`MICROSOFT_ALLOW_DEVICE_CODE=true` and `MICROSOFT_LOGIN_METHOD=device_code` are
set. Tokens are never written to plaintext files; delegated cache persistence
uses the platform keyring and otherwise remains memory-only.

For an Azure managed identity:

```text
MICROSOFT_AUTH_MODE=managed_identity
MICROSOFT_MANAGED_IDENTITY_CLIENT_ID=<optional-user-assigned-client-id>
```

For a federated workload identity:

```text
MICROSOFT_AUTH_MODE=workload_identity
MICROSOFT_TENANT_ID=<tenant-guid>
MICROSOFT_CLIENT_ID=<application-client-id>
AZURE_FEDERATED_TOKEN_FILE=/var/run/secrets/entra/tokens/identity-token
```

For certificate-backed application authentication, configure the tenant,
client ID, PEM private-key/certificate path, and thumbprint. Keep the PEM and
optional passphrase in a secret mount; do not commit them.

## Permission profiles

Profiles are additive and resolve to delegated scopes in interactive mode. The
default is `productivity,collaboration`.

| Profile | Main capability |
|---|---|
| `read_only` | Read mail, calendars, files, selected sites, chats, tasks, notes, and contacts |
| `productivity` | Send/update mail, calendars, files, tasks, notes, and contacts |
| `collaboration` | Teams chats/channels, membership, meetings, groups, and presence |
| `device_read` | Intune and Entra device inventory |
| `device_admin` | Intune updates and privileged remote device actions |
| `tenant_admin` | Directory/application/policy/security administration; never a default |

Set profiles with `MICROSOFT_PERMISSION_PROFILES`. Use
`MICROSOFT_GRAPH_SCOPES` only for a reviewed scope that is not represented by a
profile. Application and workload identities always request Graph's `.default`
scope, so their app roles must be consented in Entra.

Dataverse and the Windows companion API are separate OAuth resources. After
the normal `login` succeeds in delegated mode, call `login_power_platform` and
`login_windows_companion` once for the integrations you configured. Those
tools request consent only for resource audiences already allowlisted in the
protected integration configuration; ordinary tool calls never open an
interactive prompt or choose an arbitrary audience. Application, managed, and
workload identities instead require the corresponding application role on
each resource and continue to use non-interactive `.default` acquisition.
For the companion API, clients configure the resource as
`api://<control-plane-api-client-id>`, while the control-plane service validates
the v2 token's bare client-ID GUID `aud` claim. Configure
`api.requestedAccessTokenVersion=2` on that resource application.

The enrollment helper in `deployment/entra` resolves current permission IDs
from the tenant rather than hardcoding GUIDs. It does not create secrets or
grant admin consent. `Sites.Selected` additionally needs an explicit grant on
each SharePoint site. Dataverse and the companion control plane are separate
resource audiences and need their own API permissions/roles.

## Two independent authorization gates

Graph consent says what the identity may do. Runtime policy says what this
server is currently allowed to invoke:

```text
MICROSOFT_ALLOW_WRITES=false
MICROSOFT_ALLOW_DESTRUCTIVE=false
```

Reads work with both flags disabled. Sends, creates, edits, uploads, flow runs,
and remote actions require writes. Deletes, wipes, retirements, immediate
reboots, shutdowns, and other destructive tools require both flags. Intune and
Windows companion actions also enforce per-device/action allowlists and
time-bound confirmation evidence inside their service layer.

## Remote MCP transport

Use `stdio` for a local client. Before binding HTTP outside loopback, configure
the MCP transport's OIDC/JWT authentication, TLS at the server or reverse proxy,
an exact issuer and audience, and narrow network access. Microsoft Graph login
does not authenticate callers to the MCP server; these are separate boundaries.
