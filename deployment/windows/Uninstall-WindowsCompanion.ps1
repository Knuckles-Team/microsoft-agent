#requires -Version 5.1

<#
.SYNOPSIS
Removes the current user's Windows companion Scheduled Task and runtime.

.DESCRIPTION
Stops and unregisters only a task owned by the current user. By default the
ACL-protected configuration, data, and logs are retained. Use -PurgeData to
delete the entire per-user companion directory after making any required backup.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [Parameter()]
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "MicrosoftAgent\Companion"),

    [Parameter()]
    [ValidatePattern("^[A-Za-z0-9 ._-]{1,128}$")]
    [string]$TaskName = "Microsoft Agent Windows Companion",

    [Parameter()]
    [switch]$PurgeData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "The Windows companion can only be uninstalled on Windows."
}

$InstallRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($InstallRoot))
$statePath = Join-Path $InstallRoot "deployment-state.json"
if (-not $PSBoundParameters.ContainsKey("TaskName") -and
    (Test-Path -LiteralPath $statePath -PathType Leaf)) {
    try {
        $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (@($state.PSObject.Properties.Name) -contains "task_name") {
            $TaskName = [string]$state.task_name
            if ($TaskName -notmatch "^[A-Za-z0-9 ._-]{1,128}$") {
                throw "Deployment state contains an invalid Scheduled Task name."
            }
        }
    }
    catch {
        throw "The deployment state is invalid; specify -TaskName explicitly or repair the state file."
    }
}

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    $principal = [string]$task.Principal.UserId
    $ownedByCurrentUser = $principal -ieq $currentIdentity.Name -or
        $principal -ieq $currentIdentity.User.Value
    if (-not $ownedByCurrentUser) {
        throw "Scheduled Task '$TaskName' is not owned by the current user and will not be modified."
    }
    if ($PSCmdlet.ShouldProcess($TaskName, "Stop and unregister the per-user Scheduled Task")) {
        Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }
}

if (-not (Test-Path -LiteralPath $InstallRoot -PathType Container)) {
    Write-Output "The Windows companion is not installed at $InstallRoot."
    return
}

if ($PurgeData) {
    if ($PSCmdlet.ShouldProcess($InstallRoot, "Permanently delete runtime, configuration, data, and logs")) {
        Remove-Item -LiteralPath $InstallRoot -Recurse -Force
    }
}
else {
    foreach ($path in @(
            (Join-Path $InstallRoot "releases"),
            (Join-Path $InstallRoot "bin"),
            (Join-Path $InstallRoot "deployment-state.json")
        )) {
        if ((Test-Path -LiteralPath $path) -and $PSCmdlet.ShouldProcess($path, "Remove companion runtime component")) {
            Remove-Item -LiteralPath $path -Recurse -Force
        }
    }
    Write-Output "Configuration, data, and logs were retained under $InstallRoot."
}
