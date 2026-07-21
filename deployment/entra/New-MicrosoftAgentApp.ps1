[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [string]$DisplayName = "Microsoft Agent",

    [ValidateSet("read_only", "productivity", "collaboration", "device_read", "device_admin", "tenant_admin")]
    [string[]]$DelegatedProfiles = @("productivity", "collaboration"),

    [ValidateSet("productivity", "device_read", "device_admin")]
    [string[]]$ApplicationProfiles = @(),

    [string]$OutputPath = ".\microsoft-agent-enrollment.local.json"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$graphAppId = "00000003-0000-0000-c000-000000000000"
$profilePath = Join-Path $PSScriptRoot "permission-profiles.json"

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
    throw "Install Microsoft.Graph.Applications before running this script."
}

Import-Module Microsoft.Graph.Applications
Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All" -NoWelcome

$profiles = Get-Content -LiteralPath $profilePath -Raw | ConvertFrom-Json -AsHashtable
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property "id,appId,oauth2PermissionScopes,appRoles" | Select-Object -First 1
if (-not $graphServicePrincipal) {
    throw "Microsoft Graph service principal was not found in this tenant."
}

$resourceAccess = [System.Collections.Generic.List[object]]::new()
$resolvedDelegated = [System.Collections.Generic.List[string]]::new()
$resolvedApplication = [System.Collections.Generic.List[string]]::new()

foreach ($profile in $DelegatedProfiles) {
    foreach ($permission in $profiles.delegated[$profile]) {
        $scope = $graphServicePrincipal.Oauth2PermissionScopes | Where-Object {
            $_.Value -eq $permission -and $_.IsEnabled
        } | Select-Object -First 1
        if (-not $scope) {
            Write-Warning "Delegated permission '$permission' is unavailable and was skipped."
            continue
        }
        $resourceAccess.Add(@{ Id = $scope.Id; Type = "Scope" })
        $resolvedDelegated.Add($permission)
    }
}

foreach ($profile in $ApplicationProfiles) {
    foreach ($permission in $profiles.application[$profile]) {
        $role = $graphServicePrincipal.AppRoles | Where-Object {
            $_.Value -eq $permission -and $_.IsEnabled -and $_.AllowedMemberTypes -contains "Application"
        } | Select-Object -First 1
        if (-not $role) {
            Write-Warning "Application permission '$permission' is unavailable and was skipped."
            continue
        }
        $resourceAccess.Add(@{ Id = $role.Id; Type = "Role" })
        $resolvedApplication.Add($permission)
    }
}

$deduplicatedAccess = $resourceAccess | Sort-Object Type, Id -Unique
$body = @{
    DisplayName = $DisplayName
    SignInAudience = "AzureADMyOrg"
    IsFallbackPublicClient = $true
    PublicClient = @{
        RedirectUris = @("http://localhost")
    }
    RequiredResourceAccess = @(
        @{
            ResourceAppId = $graphAppId
            ResourceAccess = @($deduplicatedAccess)
        }
    )
}

if (-not $PSCmdlet.ShouldProcess($TenantId, "Create Entra application '$DisplayName'")) {
    return
}

$application = New-MgApplication -BodyParameter $body
$adminConsentUrl = "https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$($application.AppId)"
$result = [ordered]@{
    tenant_id = $TenantId
    client_id = $application.AppId
    object_id = $application.Id
    display_name = $application.DisplayName
    delegated_permissions = @($resolvedDelegated | Sort-Object -Unique)
    application_permissions = @($resolvedApplication | Sort-Object -Unique)
    admin_consent_url = $adminConsentUrl
    created_at_utc = [DateTimeOffset]::UtcNow.ToString("o")
}

$parent = Split-Path -Parent $OutputPath
if ($parent) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8NoBOM

Write-Host "Application created. Enrollment metadata: $OutputPath"
Write-Host "Review the requested permissions, then grant tenant admin consent at:"
Write-Host $adminConsentUrl
Write-Host "No client secret was created. Use broker login, a certificate, or a workload identity."
