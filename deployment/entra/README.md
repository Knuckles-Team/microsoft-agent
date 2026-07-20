# Entra enrollment assets

`New-MicrosoftAgentApp.ps1` creates a single-tenant public-client application
and resolves permission IDs from the tenant's live Microsoft Graph service
principal. It never creates or prints a client secret. `-WhatIf` is supported.

Run it only when authentication is ready:

```powershell
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
.\New-MicrosoftAgentApp.ps1 `
  -TenantId '<tenant-guid>' `
  -DelegatedProfiles productivity,collaboration `
  -OutputPath "$HOME\.microsoft-agent\enrollment.local.json"
```

Add `device_read` for Intune inventory. Add `device_admin` only for the remote
actions you intend to allow. `tenant_admin` is intentionally excluded from the
script's defaults. The script writes local enrollment metadata to a filename
matched by `.gitignore`; it does not grant consent automatically.

For application mode, request only the `ApplicationProfiles` required by a
daemon, grant admin consent, and use a certificate, managed identity, or
federated workload identity. Teams message sending is designed primarily for a
delegated user in this project and is not included in application profiles.

`Sites.Selected` also requires a separate site-level permission grant for each
SharePoint site. A tenant-wide consent grant alone does not authorize any site.

## Windows control-plane identities

Use `New-WindowsControlPlaneEntraResources.ps1` after creating the main
Microsoft Agent application above. It creates a separate single-tenant v2 API
resource, exposes `WindowsCompanion.Control` as both an admin-restricted
delegated scope and application role, exposes `WindowsCompanion.Device` only as
an application role, and uploads a public certificate to a separate worker
application. It never creates a client secret or uploads a private key.

Export the worker authentication certificate as a public binary DER `.cer`
file. Keep its private key on the enrolled laptop. Preview the full operation
without connecting to Entra:

```powershell
.\New-WindowsControlPlaneEntraResources.ps1 `
  -TenantId '<tenant-guid>' `
  -ControllerApplicationId '<microsoft-agent-client-id>' `
  -WorkerCertificatePath 'C:\SecureStaging\laptop-01-public.cer' `
  -WorkerDisplayName 'Microsoft Agent Worker - Laptop 01' `
  -WhatIf
```

Run without `-WhatIf` after review. By default the helper only declares
`requiredResourceAccess`; it makes no permission grant. Review the emitted
`.local.json` metadata and use the recorded admin-consent URLs, or rerun with
`-GrantAdminConsent` only when the signed-in administrator is authorized to
grant both the controller permission and worker application role:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser

.\New-WindowsControlPlaneEntraResources.ps1 `
  -TenantId '<tenant-guid>' `
  -ControllerApplicationId '<microsoft-agent-client-id>' `
  -WorkerCertificatePath 'C:\SecureStaging\laptop-01-public.cer' `
  -WorkerDisplayName 'Microsoft Agent Worker - Laptop 01' `
  -OutputPath "$HOME\.microsoft-agent\laptop-01-enrollment.local.control-plane.json"
```

Use `-ControllerPermissionMode Application` for a daemon controller or `Both`
only when both interactive and daemon controller modes are required. On a
rerun, pass the emitted `-ApiApplicationId` and `-WorkerApplicationId`; the
script validates the existing API contract, merges missing permission
declarations, avoids duplicate role grants, and avoids re-uploading the same
certificate. Changes are additive: the helper never removes an existing
permission or credential. Run it separately for every worker identity and
certificate.

The helper requests `Application.ReadWrite.All`. Only the explicit
`-GrantAdminConsent` path additionally requests
`AppRoleAssignment.ReadWrite.All` and
`DelegatedPermissionGrant.ReadWrite.All`. The administrator must still hold an
appropriate Entra directory role.
