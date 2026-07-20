#Requires -Version 7.2

<#
.SYNOPSIS
Creates the separate Entra API and worker identity used by the Windows control plane.

.DESCRIPTION
Creates or validates a single-tenant v2 resource application that exposes the
WindowsCompanion.Control delegated scope and application role, plus the
WindowsCompanion.Device application role. It declares the selected controller
permission, creates or updates a separate worker application, and uploads only
the worker's public DER certificate.

No client secret is created. Application-role assignments and tenant-wide
delegated consent are made only when -GrantAdminConsent is supplied.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")]
    [string]$ControllerApplicationId,

    [ValidateSet("Delegated", "Application", "Both")]
    [string]$ControllerPermissionMode = "Delegated",

    [ValidateLength(1, 120)]
    [string]$ApiDisplayName = "Microsoft Agent Windows Control Plane API",

    [ValidatePattern("^$|^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")]
    [string]$ApiApplicationId = "",

    [ValidateLength(1, 120)]
    [string]$WorkerDisplayName = "Microsoft Agent Windows Companion Worker",

    [ValidatePattern("^$|^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")]
    [string]$WorkerApplicationId = "",

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkerCertificatePath,

    [ValidateLength(1, 90)]
    [string]$WorkerCertificateDisplayName = "Windows companion authentication",

    [switch]$GrantAdminConsent,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = ".\windows-control-plane-enrollment.local.control-plane.json"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$controlPermission = "WindowsCompanion.Control"
$devicePermission = "WindowsCompanion.Device"

function Get-ApplicationByClientId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [string[]]$Property = @(
            "id",
            "appId",
            "displayName",
            "api",
            "appRoles",
            "identifierUris",
            "requiredResourceAccess",
            "keyCredentials"
        )
    )

    $applications = @(
        Get-MgApplication -Filter "appId eq '$ClientId'" -Property $Property
    )
    if ($applications.Count -ne 1) {
        throw "Expected one application with client ID '$ClientId'; found $($applications.Count)."
    }
    return $applications[0]
}

function Get-OrCreateServicePrincipal {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId
    )

    $servicePrincipal = @(
        Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -Property "id,appId,displayName"
    ) | Select-Object -First 1
    if ($servicePrincipal) {
        return $servicePrincipal
    }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            return New-MgServicePrincipal -AppId $ClientId
        }
        catch {
            if ($attempt -eq 6) {
                throw
            }
            Start-Sleep -Seconds (2 * $attempt)
        }
    }
    throw "Service-principal creation did not complete."
}

function Set-ApiRequiredResourceAccess {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,

        [Parameter(Mandatory = $true)]
        [string]$ResourceClientId,

        [Parameter(Mandatory = $true)]
        [object[]]$RequestedAccess
    )

    $updated = $false
    $requirements = [System.Collections.Generic.List[object]]::new()
    $resourceRequirement = $null

    foreach ($requirement in @($Application.RequiredResourceAccess)) {
        $access = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in @($requirement.ResourceAccess)) {
            $access.Add(@{
                Id = [Guid]$entry.Id
                Type = [string]$entry.Type
            })
        }
        $converted = @{
            ResourceAppId = [string]$requirement.ResourceAppId
            ResourceAccess = $access.ToArray()
        }
        $requirements.Add($converted)
        if ([string]$requirement.ResourceAppId -eq $ResourceClientId) {
            $resourceRequirement = $converted
        }
    }

    if (-not $resourceRequirement) {
        $resourceRequirement = @{
            ResourceAppId = $ResourceClientId
            ResourceAccess = @()
        }
        $requirements.Add($resourceRequirement)
        $updated = $true
    }

    $resourceEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($resourceRequirement.ResourceAccess)) {
        $resourceEntries.Add($entry)
    }
    foreach ($entry in $RequestedAccess) {
        $alreadyPresent = @($resourceEntries) | Where-Object {
            [string]$_.Id -eq [string]$entry.Id -and
            [string]$_.Type -eq [string]$entry.Type
        } | Select-Object -First 1
        if (-not $alreadyPresent) {
            $resourceEntries.Add($entry)
            $updated = $true
        }
    }
    $resourceRequirement.ResourceAccess = $resourceEntries.ToArray()

    if ($updated) {
        Update-MgApplication `
            -ApplicationId $Application.Id `
            -RequiredResourceAccess $requirements.ToArray()
    }
}

function Grant-ApplicationRole {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ClientServicePrincipal,

        [Parameter(Mandatory = $true)]
        [object]$ResourceServicePrincipal,

        [Parameter(Mandatory = $true)]
        [Guid]$RoleId
    )

    $existing = @(
        Get-MgServicePrincipalAppRoleAssignedTo `
            -ServicePrincipalId $ResourceServicePrincipal.Id `
            -All
    ) | Where-Object {
        [string]$_.PrincipalId -eq [string]$ClientServicePrincipal.Id -and
        [string]$_.AppRoleId -eq [string]$RoleId
    } | Select-Object -First 1

    if (-not $existing) {
        New-MgServicePrincipalAppRoleAssignedTo `
            -ServicePrincipalId $ResourceServicePrincipal.Id `
            -BodyParameter @{
                PrincipalId = $ClientServicePrincipal.Id
                ResourceId = $ResourceServicePrincipal.Id
                AppRoleId = $RoleId
            } | Out-Null
    }
}

function Grant-TenantWideDelegatedScope {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ClientServicePrincipal,

        [Parameter(Mandatory = $true)]
        [object]$ResourceServicePrincipal,

        [Parameter(Mandatory = $true)]
        [string]$Scope
    )

    $filter = "clientId eq '$($ClientServicePrincipal.Id)' and resourceId eq '$($ResourceServicePrincipal.Id)' and consentType eq 'AllPrincipals'"
    $grant = @(
        Get-MgOauth2PermissionGrant -Filter $filter -All
    ) | Select-Object -First 1
    if (-not $grant) {
        New-MgOauth2PermissionGrant -BodyParameter @{
            ClientId = $ClientServicePrincipal.Id
            ConsentType = "AllPrincipals"
            ResourceId = $ResourceServicePrincipal.Id
            Scope = $Scope
        } | Out-Null
        return
    }

    $scopes = @(
        ([string]$grant.Scope).Split(
            " ",
            [System.StringSplitOptions]::RemoveEmptyEntries
        )
    )
    if ($scopes -notcontains $Scope) {
        $scopes = @($scopes + $Scope | Sort-Object -Unique)
        Update-MgOauth2PermissionGrant `
            -OAuth2PermissionGrantId $grant.Id `
            -BodyParameter @{ Scope = $scopes -join " " }
    }
}

$certificateFile = Get-Item -LiteralPath $WorkerCertificatePath
if ($certificateFile.PSIsContainer -or $certificateFile.LinkType) {
    throw "WorkerCertificatePath must be a regular, non-link DER .cer file."
}
if ($certificateFile.Extension -ne ".cer") {
    throw "WorkerCertificatePath must use a public DER .cer certificate."
}

$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    [System.IO.File]::ReadAllBytes($certificateFile.FullName)
)
if ($certificate.HasPrivateKey) {
    throw "The worker certificate file must not contain a private key."
}
$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey(
    $certificate
)
if (-not $rsa -or $rsa.KeySize -lt 2048) {
    throw "The worker certificate must contain an RSA public key of at least 2048 bits."
}
$rsa.Dispose()
if ($certificate.NotBefore.ToUniversalTime() -gt [DateTime]::UtcNow.AddMinutes(5)) {
    throw "The worker certificate is not valid yet."
}
if ($certificate.NotAfter.ToUniversalTime() -le [DateTime]::UtcNow) {
    throw "The worker certificate has expired."
}
if ($certificate.NotAfter.ToUniversalTime() -lt [DateTime]::UtcNow.AddDays(30)) {
    Write-Warning "The worker certificate expires in less than 30 days."
}

$operation = "Provision a v2 control-plane API, declare controller/worker permissions, and upload one public certificate"
if ($GrantAdminConsent) {
    $operation += "; grant tenant-wide delegated/application consent"
}
if (-not $PSCmdlet.ShouldProcess($TenantId, $operation)) {
    $certificate.Dispose()
    return
}

$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Applications"
)
if ($GrantAdminConsent) {
    $requiredModules += "Microsoft.Graph.Identity.SignIns"
}
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw "Install the '$module' module before running this script."
    }
    Import-Module $module
}

$connectScopes = [System.Collections.Generic.List[string]]::new()
$connectScopes.Add("Application.ReadWrite.All")
if ($GrantAdminConsent) {
    $connectScopes.Add("AppRoleAssignment.ReadWrite.All")
    $connectScopes.Add("DelegatedPermissionGrant.ReadWrite.All")
}
Connect-MgGraph -TenantId $TenantId -Scopes $connectScopes.ToArray() -NoWelcome

$controllerApplication = Get-ApplicationByClientId -ClientId $ControllerApplicationId

if ($ApiApplicationId) {
    $apiApplication = Get-ApplicationByClientId -ClientId $ApiApplicationId
    if ([int]$apiApplication.Api.RequestedAccessTokenVersion -ne 2) {
        throw "The existing API application must set api.requestedAccessTokenVersion to 2."
    }
    if (@($apiApplication.IdentifierUris) -notcontains "api://$ApiApplicationId") {
        throw "The existing API application must expose the identifier URI api://$ApiApplicationId."
    }
    $controlScope = @($apiApplication.Api.Oauth2PermissionScopes) | Where-Object {
        $_.Value -eq $controlPermission -and $_.IsEnabled
    } | Select-Object -First 1
    $controlRole = @($apiApplication.AppRoles) | Where-Object {
        $_.Value -eq $controlPermission -and
        $_.IsEnabled -and
        $_.AllowedMemberTypes -contains "Application"
    } | Select-Object -First 1
    $deviceRole = @($apiApplication.AppRoles) | Where-Object {
        $_.Value -eq $devicePermission -and
        $_.IsEnabled -and
        $_.AllowedMemberTypes -contains "Application"
    } | Select-Object -First 1
    if (-not $controlScope -or -not $controlRole -or -not $deviceRole) {
        throw "The existing API is missing one or more required enabled scope/role definitions."
    }
}
else {
    $controlScopeId = [Guid]::NewGuid()
    $controlRoleId = [Guid]::NewGuid()
    $deviceRoleId = [Guid]::NewGuid()
    $apiApplication = New-MgApplication -BodyParameter @{
        DisplayName = $ApiDisplayName
        SignInAudience = "AzureADMyOrg"
        Api = @{
            RequestedAccessTokenVersion = 2
            Oauth2PermissionScopes = @(
                @{
                    Id = $controlScopeId
                    IsEnabled = $true
                    Type = "Admin"
                    Value = $controlPermission
                    AdminConsentDisplayName = "Control enrolled Windows companions"
                    AdminConsentDescription = "Submit and inspect typed actions for explicitly enrolled Windows companion devices."
                }
            )
        }
        AppRoles = @(
            @{
                Id = $controlRoleId
                IsEnabled = $true
                AllowedMemberTypes = @("Application")
                Value = $controlPermission
                DisplayName = "Control enrolled Windows companions"
                Description = "Allows a daemon controller to submit and inspect typed actions for enrolled devices."
            },
            @{
                Id = $deviceRoleId
                IsEnabled = $true
                AllowedMemberTypes = @("Application")
                Value = $devicePermission
                DisplayName = "Poll as an enrolled Windows companion"
                Description = "Allows a certificate-backed worker to poll and acknowledge its device queue."
            }
        )
    }
    Update-MgApplication `
        -ApplicationId $apiApplication.Id `
        -IdentifierUris @("api://$($apiApplication.AppId)")
    $apiApplication = Get-ApplicationByClientId -ClientId $apiApplication.AppId
    $controlScope = @($apiApplication.Api.Oauth2PermissionScopes) | Where-Object {
        $_.Value -eq $controlPermission
    } | Select-Object -First 1
    $controlRole = @($apiApplication.AppRoles) | Where-Object {
        $_.Value -eq $controlPermission
    } | Select-Object -First 1
    $deviceRole = @($apiApplication.AppRoles) | Where-Object {
        $_.Value -eq $devicePermission
    } | Select-Object -First 1
}

if ($apiApplication.AppId -eq $controllerApplication.AppId) {
    throw "The control-plane API and controller must be separate applications."
}

$controllerAccess = [System.Collections.Generic.List[object]]::new()
if ($ControllerPermissionMode -in @("Delegated", "Both")) {
    $controllerAccess.Add(@{ Id = [Guid]$controlScope.Id; Type = "Scope" })
}
if ($ControllerPermissionMode -in @("Application", "Both")) {
    $controllerAccess.Add(@{ Id = [Guid]$controlRole.Id; Type = "Role" })
}
Set-ApiRequiredResourceAccess `
    -Application $controllerApplication `
    -ResourceClientId $apiApplication.AppId `
    -RequestedAccess $controllerAccess.ToArray()

if ($WorkerApplicationId) {
    $workerApplication = Get-ApplicationByClientId -ClientId $WorkerApplicationId
}
else {
    $workerApplication = New-MgApplication -BodyParameter @{
        DisplayName = $WorkerDisplayName
        SignInAudience = "AzureADMyOrg"
        RequiredResourceAccess = @(
            @{
                ResourceAppId = $apiApplication.AppId
                ResourceAccess = @(
                    @{ Id = [Guid]$deviceRole.Id; Type = "Role" }
                )
            }
        )
    }
    $workerApplication = Get-ApplicationByClientId -ClientId $workerApplication.AppId
}

if ($workerApplication.AppId -in @(
    $apiApplication.AppId,
    $controllerApplication.AppId
)) {
    throw "The worker, controller, and control-plane API must be separate applications."
}

Set-ApiRequiredResourceAccess `
    -Application $workerApplication `
    -ResourceClientId $apiApplication.AppId `
    -RequestedAccess @(@{ Id = [Guid]$deviceRole.Id; Type = "Role" })

$workerApplication = Get-ApplicationByClientId -ClientId $workerApplication.AppId
$certificateIdentifier = [Convert]::ToBase64String($certificate.GetCertHash())
$existingCertificate = @($workerApplication.KeyCredentials) | Where-Object {
    $_.CustomKeyIdentifier -and
    [Convert]::ToBase64String($_.CustomKeyIdentifier) -eq $certificateIdentifier
} | Select-Object -First 1
if (-not $existingCertificate) {
    $keyCredentials = [System.Collections.Generic.List[object]]::new()
    foreach ($credential in @($workerApplication.KeyCredentials)) {
        $keyCredentials.Add($credential)
    }
    $keyCredentials.Add(@{
        CustomKeyIdentifier = $certificate.GetCertHash()
        DisplayName = $WorkerCertificateDisplayName
        EndDateTime = $certificate.NotAfter.ToUniversalTime()
        Key = $certificate.RawData
        KeyId = [Guid]::NewGuid()
        StartDateTime = $certificate.NotBefore.ToUniversalTime()
        Type = "AsymmetricX509Cert"
        Usage = "Verify"
    })
    Update-MgApplication `
        -ApplicationId $workerApplication.Id `
        -KeyCredentials $keyCredentials.ToArray()
}

$apiServicePrincipal = Get-OrCreateServicePrincipal -ClientId $apiApplication.AppId
$controllerServicePrincipal = Get-OrCreateServicePrincipal `
    -ClientId $controllerApplication.AppId
$workerServicePrincipal = Get-OrCreateServicePrincipal `
    -ClientId $workerApplication.AppId

if ($GrantAdminConsent) {
    if ($ControllerPermissionMode -in @("Delegated", "Both")) {
        Grant-TenantWideDelegatedScope `
            -ClientServicePrincipal $controllerServicePrincipal `
            -ResourceServicePrincipal $apiServicePrincipal `
            -Scope $controlPermission
    }
    if ($ControllerPermissionMode -in @("Application", "Both")) {
        Grant-ApplicationRole `
            -ClientServicePrincipal $controllerServicePrincipal `
            -ResourceServicePrincipal $apiServicePrincipal `
            -RoleId ([Guid]$controlRole.Id)
    }
    Grant-ApplicationRole `
        -ClientServicePrincipal $workerServicePrincipal `
        -ResourceServicePrincipal $apiServicePrincipal `
        -RoleId ([Guid]$deviceRole.Id)
}

$result = [ordered]@{
    tenant_id = $TenantId
    api = [ordered]@{
        client_id = $apiApplication.AppId
        object_id = $apiApplication.Id
        service_principal_object_id = $apiServicePrincipal.Id
        identifier_uri = "api://$($apiApplication.AppId)"
        token_audience = $apiApplication.AppId
        requested_access_token_version = 2
        controller_scope_id = $controlScope.Id
        controller_role_id = $controlRole.Id
        device_role_id = $deviceRole.Id
    }
    controller = [ordered]@{
        client_id = $controllerApplication.AppId
        object_id = $controllerApplication.Id
        service_principal_object_id = $controllerServicePrincipal.Id
        permission_mode = $ControllerPermissionMode
    }
    worker = [ordered]@{
        client_id = $workerApplication.AppId
        object_id = $workerApplication.Id
        service_principal_object_id = $workerServicePrincipal.Id
        certificate_thumbprint = $certificate.Thumbprint.ToUpperInvariant()
        certificate_not_after_utc = $certificate.NotAfter.ToUniversalTime().ToString("o")
    }
    admin_consent_granted_by_script = [bool]$GrantAdminConsent
    controller_admin_consent_url = "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$($controllerApplication.AppId)"
    worker_admin_consent_url = "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$($workerApplication.AppId)"
    created_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
}

$parent = Split-Path -Parent $OutputPath
if ($parent) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8NoBOM
$certificate.Dispose()

Write-Host "Control-plane enrollment metadata: $OutputPath"
Write-Host "No client secret or private key was created or uploaded."
if (-not $GrantAdminConsent) {
    Write-Host "No admin consent was granted. Review requiredResourceAccess, then use the recorded consent URLs or rerun with -GrantAdminConsent."
}
