#requires -Version 5.1

<#
.SYNOPSIS
Runs the installed outbound Windows companion from its versioned environment.

.DESCRIPTION
This wrapper is the Scheduled Task target. It reads only path metadata from the
ACL-protected deployment state, takes a per-user mutex, performs bounded log
retention, and launches the fixed companion module. It never accepts tokens,
passwords, commands, URLs, or arbitrary Python modules.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$StatePath = (Join-Path $env:LOCALAPPDATA "MicrosoftAgent\Companion\deployment-state.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-DeploymentState {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
        throw "The companion deployment state does not exist. Run Install-WindowsCompanion.ps1 first."
    }
    try {
        $value = Get-Content -LiteralPath $LiteralPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "The companion deployment state is not valid JSON."
    }
    foreach ($property in @("current_release", "config_path", "data_root", "log_root")) {
        if (-not $value.$property -or $value.$property -isnot [string]) {
            throw "The companion deployment state is missing '$property'."
        }
    }
    if ($value.python_module -and $value.python_module -ne "microsoft_agent.windows_companion_service") {
        throw "The deployment state requested an unsupported Python module."
    }
    return $value
}

function Remove-ExpiredLogs {
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter()][int]$Keep = 20
    )

    Get-ChildItem -LiteralPath $LiteralPath -File -Filter "companion-*.log" -ErrorAction SilentlyContinue |
        Sort-Object -Property LastWriteTimeUtc -Descending |
        Select-Object -Skip $Keep |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw "The Windows companion runtime can only run on Windows."
}

$StatePath = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($StatePath))
$state = Get-DeploymentState -LiteralPath $StatePath
$releasePath = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables([string]$state.current_release))
$pythonPath = Join-Path $releasePath "venv\Scripts\python.exe"
$configPath = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables([string]$state.config_path))
$dataRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables([string]$state.data_root))
$logRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables([string]$state.log_root))

foreach ($file in @($pythonPath, $configPath)) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "A required companion file is missing: $file"
    }
}
foreach ($directory in @($dataRoot, $logRoot)) {
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        throw "A required companion directory is missing: $directory"
    }
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$mutexName = "Local\MicrosoftAgentWindowsCompanion-$($identity.User.Value)"
$mutex = [Threading.Mutex]::new($false, $mutexName)
$hasMutex = $false
try {
    $hasMutex = $mutex.WaitOne(0)
    if (-not $hasMutex) {
        Write-Output "The companion is already running for this user."
        exit 0
    }

    Remove-ExpiredLogs -LiteralPath $logRoot
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $stdoutPath = Join-Path $logRoot "companion-$timestamp-stdout.log"
    $stderrPath = Join-Path $logRoot "companion-$timestamp-stderr.log"

    # These variables contain paths only. Authentication secrets are never
    # passed on the command line or written by the deployment wrapper.
    $env:MICROSOFT_INTEGRATIONS_CONFIG_PATH = $configPath
    $env:MICROSOFT_WINDOWS_COMPANION_DATA_ROOT = $dataRoot

    & $pythonPath -m microsoft_agent.windows_companion_service --config $configPath `
        1>> $stdoutPath 2>> $stderrPath
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        Write-Error "The Windows companion stopped with exit code $exitCode. See the protected log directory."
    }
    exit $exitCode
}
finally {
    if ($hasMutex) {
        $mutex.ReleaseMutex()
    }
    $mutex.Dispose()
}
