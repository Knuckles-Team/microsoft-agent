#requires -Version 5.1

<#
.SYNOPSIS
Performs Pester-style, read-only validation of a companion deployment.

.DESCRIPTION
Checks deployment state, Python/package health, service configuration, ACLs,
and Scheduled Task isolation. No authentication request or network call is made.
The script exits 0 when no check failed and 1 otherwise.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "MicrosoftAgent\Companion"),

    [Parameter()]
    [ValidatePattern("^[A-Za-z0-9 ._-]{1,128}$")]
    [string]$TaskName = "Microsoft Agent Windows Companion",

    [Parameter()]
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$results = [Collections.Generic.List[object]]::new()

function Add-Result {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet("Pass", "Warn", "Fail")][string]$Status,
        [Parameter(Mandatory = $true)][string]$Message
    )
    $script:results.Add([PSCustomObject]@{
            Name    = $Name
            Status  = $Status
            Message = $Message
        })
}

function Test-ProtectedAcl {
    param(
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string[]]$AllowedSids
    )

    $acl = Get-Acl -LiteralPath $LiteralPath
    if (-not $acl.AreAccessRulesProtected) {
        throw "ACL inheritance is enabled."
    }
    foreach ($rule in $acl.Access) {
        try {
            $sid = $rule.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value
        }
        catch {
            throw "ACL contains an unresolvable identity."
        }
        if ($sid -notin $AllowedSids) {
            throw "ACL grants access to unexpected identity $sid."
        }
    }
}

if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    Add-Result -Name "Windows host" -Status "Fail" -Message "Validation must run on Windows."
}
else {
    Add-Result -Name "Windows host" -Status "Pass" -Message "Running on Windows with PowerShell $($PSVersionTable.PSVersion)."
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$allowedSids = @($identity.User.Value, "S-1-5-18", "S-1-5-32-544")
$InstallRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($InstallRoot))
$statePath = Join-Path $InstallRoot "deployment-state.json"
$state = $null

try {
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
        throw "Deployment state was not found at $statePath."
    }
    $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($state.schema_version -ne 1) {
        throw "Unsupported deployment state schema."
    }
    if ($state.python_module -ne "microsoft_agent.windows_companion_service") {
        throw "Deployment state references an unexpected module."
    }
    Add-Result -Name "Deployment state" -Status "Pass" -Message "Deployment state schema and module are valid."
}
catch {
    Add-Result -Name "Deployment state" -Status "Fail" -Message $_.Exception.Message
}

foreach ($path in @(
        $InstallRoot,
        (Join-Path $InstallRoot "config"),
        (Join-Path $InstallRoot "data"),
        (Join-Path $InstallRoot "logs")
    )) {
    try {
        if (-not (Test-Path -LiteralPath $path -PathType Container)) {
            throw "Directory does not exist."
        }
        Test-ProtectedAcl -LiteralPath $path -AllowedSids $allowedSids
        Add-Result -Name "Protected directory: $path" -Status "Pass" -Message "ACL inheritance is disabled and principals are restricted."
    }
    catch {
        Add-Result -Name "Protected directory: $path" -Status "Fail" -Message $_.Exception.Message
    }
}

if ($state) {
    $pythonPath = Join-Path ([string]$state.current_release) "venv\Scripts\python.exe"
    try {
        if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
            throw "The active virtual-environment Python executable is missing."
        }
        $versionText = & $pythonPath -c "import sys; print('.'.join(map(str, sys.version_info[:3])))"
        if ($LASTEXITCODE -ne 0) {
            throw "The active Python executable failed."
        }
        $version = [Version]($versionText | Select-Object -Last 1)
        if ($version -lt [Version]"3.11" -or $version -ge [Version]"3.15") {
            throw "Python $version is outside the supported 3.11 through 3.14 range."
        }
        Add-Result -Name "Python runtime" -Status "Pass" -Message "Active release uses Python $version."
    }
    catch {
        Add-Result -Name "Python runtime" -Status "Fail" -Message $_.Exception.Message
    }

    try {
        $help = & $pythonPath -m microsoft_agent.windows_companion_service --help 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "The installed companion service entry point did not load."
        }
        Add-Result -Name "Companion module" -Status "Pass" -Message "The fixed service module and its dependencies load successfully."

        $configPath = [string]$state.config_path
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
            throw "The installed configuration file is missing."
        }
        Test-ProtectedAcl -LiteralPath $configPath -AllowedSids $allowedSids
        $null = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (($help -join "`n") -match "--validate-config") {
            & $pythonPath -m microsoft_agent.windows_companion_service `
                --config $configPath --validate-config *> $null
            if ($LASTEXITCODE -ne 0) {
                throw "The companion rejected its configuration."
            }
        }
        Add-Result -Name "Companion configuration" -Status "Pass" -Message "Configuration is valid JSON, protected, and accepted by available validation."
    }
    catch {
        Add-Result -Name "Companion configuration" -Status "Fail" -Message $_.Exception.Message
    }
}

try {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    $principal = [string]$task.Principal.UserId
    if ($principal -ine $identity.Name -and $principal -ine $identity.User.Value) {
        throw "Task principal is not the current user."
    }
    if ([string]$task.Principal.LogonType -notmatch "Interactive") {
        throw "Task does not use an interactive-user logon token."
    }
    if ([string]$task.Principal.RunLevel -notmatch "Limited") {
        throw "Task is configured to run elevated."
    }
    $actionText = (($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join " ")
    if ($actionText -notmatch "Run-WindowsCompanion\.ps1" -or
        $actionText -match "ExecutionPolicy\s+(Bypass|Unrestricted)") {
        throw "Task action is unexpected or weakens PowerShell execution policy."
    }
    if ([int]$task.Settings.RestartCount -lt 1) {
        throw "Task retry policy is disabled."
    }
    $hasLogonTrigger = $false
    foreach ($trigger in $task.Triggers) {
        if ($trigger.CimClass.CimClassName -eq "MSFT_TaskLogonTrigger") {
            $hasLogonTrigger = $true
        }
    }
    if (-not $hasLogonTrigger) {
        throw "Task is missing its per-user logon trigger."
    }
    Add-Result -Name "Scheduled Task" -Status "Pass" -Message "Task is limited to the current interactive user with logon and retry behavior."

    $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName
    if ($task.State -eq "Running") {
        Add-Result -Name "Runtime health" -Status "Pass" -Message "The outbound companion task is running."
    }
    elseif ($taskInfo.LastRunTime -eq [DateTime]::MinValue) {
        Add-Result -Name "Runtime health" -Status "Warn" -Message "The task has not run yet; this is expected while the sample device remains disabled."
    }
    else {
        Add-Result -Name "Runtime health" -Status "Warn" -Message "Task state is $($task.State), last result is $($taskInfo.LastTaskResult). Check protected logs."
    }
}
catch {
    Add-Result -Name "Scheduled Task" -Status "Fail" -Message $_.Exception.Message
}

if ($AsJson) {
    $results | ConvertTo-Json -Depth 4
}
else {
    $results | Format-Table -AutoSize
}

if ($results.Where({ $_.Status -eq "Fail" }).Count -gt 0) {
    exit 1
}
exit 0
