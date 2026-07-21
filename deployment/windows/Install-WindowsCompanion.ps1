#requires -Version 5.1

<#
.SYNOPSIS
Installs the Microsoft Agent Windows companion for the current interactive user.

.DESCRIPTION
Creates a versioned virtual environment, installs the local microsoft-agent
package with its Windows integration extras, protects the companion directories,
and registers a per-user Scheduled Task. The task runs only in the interactive
user session so Office COM, clipboard, and notification adapters work correctly.

No secret value is accepted as a parameter. Authentication material is loaded
by the companion from paths referenced in its ACL-protected configuration.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter()]
    [string]$PackageSource,

    [Parameter()]
    [string]$ConfigPath,

    [Parameter()]
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "MicrosoftAgent\Companion"),

    [Parameter()]
    [string]$PythonCommand = "py.exe",

    [Parameter()]
    [ValidatePattern("^[A-Za-z0-9 ._-]{1,128}$")]
    [string]$TaskName = "Microsoft Agent Windows Companion",

    [Parameter()]
    [switch]$ReplaceConfig,

    [Parameter()]
    [switch]$NoStart,

    [Parameter()]
    [switch]$Rollback
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-WindowsHost {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw "The Windows companion can only be installed on Windows."
    }
    if (-not $env:LOCALAPPDATA) {
        throw "LOCALAPPDATA is unavailable. Run the installer in an interactive user session."
    }
}

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    try {
        return (Resolve-Path -LiteralPath $LiteralPath -ErrorAction Stop).ProviderPath
    }
    catch {
        throw "$Description does not exist or cannot be accessed: $LiteralPath"
    }
}

function Get-CurrentIdentity {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $identity.User) {
        throw "The current Windows security identifier could not be determined."
    }
    return $identity
}

function Protect-Directory {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $acl = [Security.AccessControl.DirectorySecurity]::new()
    $acl.SetOwner($script:CurrentIdentity.User)
    $acl.SetAccessRuleProtection($true, $false)
    $inheritance = [Security.AccessControl.InheritanceFlags]::ContainerInherit -bor
        [Security.AccessControl.InheritanceFlags]::ObjectInherit
    $propagation = [Security.AccessControl.PropagationFlags]::None
    $allow = [Security.AccessControl.AccessControlType]::Allow
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            $script:CurrentIdentity.User,
            [Security.AccessControl.FileSystemRights]::FullControl,
            $inheritance,
            $propagation,
            $allow
        ))
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            [Security.Principal.SecurityIdentifier]::new("S-1-5-18"),
            [Security.AccessControl.FileSystemRights]::FullControl,
            $inheritance,
            $propagation,
            $allow
        ))
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            [Security.Principal.SecurityIdentifier]::new("S-1-5-32-544"),
            [Security.AccessControl.FileSystemRights]::FullControl,
            $inheritance,
            $propagation,
            $allow
        ))
    [IO.Directory]::SetAccessControl($LiteralPath, $acl)
}

function Protect-File {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $acl = [Security.AccessControl.FileSecurity]::new()
    $acl.SetOwner($script:CurrentIdentity.User)
    $acl.SetAccessRuleProtection($true, $false)
    $allow = [Security.AccessControl.AccessControlType]::Allow
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            $script:CurrentIdentity.User,
            [Security.AccessControl.FileSystemRights]::FullControl,
            $allow
        ))
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            [Security.Principal.SecurityIdentifier]::new("S-1-5-18"),
            [Security.AccessControl.FileSystemRights]::FullControl,
            $allow
        ))
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            [Security.Principal.SecurityIdentifier]::new("S-1-5-32-544"),
            [Security.AccessControl.FileSystemRights]::FullControl,
            $allow
        ))
    [IO.File]::SetAccessControl($LiteralPath, $acl)
}

function Invoke-SelectedPython {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    if ($script:UsePythonLauncher) {
        & $script:PythonExecutable -3 @Arguments
    }
    else {
        & $script:PythonExecutable @Arguments
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Python failed with exit code $LASTEXITCODE."
    }
}

function Test-PythonPrerequisite {
    $command = Get-Command -Name $PythonCommand -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Python was not found using '$PythonCommand'. Install 64-bit Python 3.11 through 3.14 first."
    }
    $script:PythonExecutable = $command.Source
    $script:UsePythonLauncher = [IO.Path]::GetFileName($command.Source) -ieq "py.exe"

    if ($script:UsePythonLauncher) {
        $version = & $script:PythonExecutable -3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    }
    else {
        $version = & $script:PythonExecutable -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
    }
    if ($LASTEXITCODE -ne 0 -or -not $version) {
        throw "Python could not be executed using '$PythonCommand'."
    }
    $parsed = [Version]($version | Select-Object -Last 1)
    if ($parsed -lt [Version]"3.11" -or $parsed -ge [Version]"3.15") {
        throw "Python $parsed is unsupported. Install Python 3.11 through 3.14."
    }
    return $parsed
}

function Read-State {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
        return $null
    }
    try {
        return Get-Content -LiteralPath $LiteralPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "The deployment state file is invalid: $LiteralPath"
    }
}

function Write-State {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$LiteralPath
    )

    $temporary = "$LiteralPath.new"
    $json = $State | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($temporary, $json, [Text.UTF8Encoding]::new($false))
    Protect-File -LiteralPath $temporary
    Move-Item -LiteralPath $temporary -Destination $LiteralPath -Force
    Protect-File -LiteralPath $LiteralPath
}

function Register-CompanionTask {
    param(
        [Parameter(Mandatory = $true)][string]$RunnerPath,
        [Parameter(Mandatory = $true)][string]$StatePath
    )

    $powerShell = Join-Path $PSHOME "powershell.exe"
    if (-not (Test-Path -LiteralPath $powerShell -PathType Leaf)) {
        $powerShell = (Get-Command -Name "powershell.exe" -ErrorAction Stop).Source
    }
    $arguments = '-NoLogo -NoProfile -NonInteractive -File "{0}" -StatePath "{1}"' -f `
        $RunnerPath.Replace('"', '""'), $StatePath.Replace('"', '""')
    $action = New-ScheduledTaskAction -Execute $powerShell -Argument $arguments -WorkingDirectory $InstallRoot
    $logonTrigger = New-ScheduledTaskTrigger -AtLogOn -User $script:CurrentIdentity.Name
    $retryTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) `
        -RepetitionInterval (New-TimeSpan -Minutes 15) `
        -RepetitionDuration (New-TimeSpan -Days 3650)
    $principal = New-ScheduledTaskPrincipal -UserId $script:CurrentIdentity.Name `
        -LogonType Interactive -RunLevel Limited
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable -MultipleInstances IgnoreNew `
        -RestartCount 5 -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit ([TimeSpan]::Zero)
    Register-ScheduledTask -TaskName $TaskName -Action $action `
        -Trigger @($logonTrigger, $retryTrigger) -Principal $principal `
        -Settings $settings -Description `
        "Outbound-only Microsoft Agent companion for the current interactive user." `
        -Force | Out-Null
}

function Test-ConfigurationReadiness {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    try {
        $config = Get-Content -LiteralPath $LiteralPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "The companion configuration is not valid JSON: $LiteralPath"
    }
    $propertyNames = @($config.PSObject.Properties.Name)
    $hasDirectConfig = $propertyNames -contains "control_plane_url"
    $hasIntegratedConfig = $propertyNames -contains "windows_companion"
    if (-not $hasDirectConfig -and -not $hasIntegratedConfig) {
        throw "The companion configuration does not contain control_plane_url or windows_companion settings."
    }
    if (($propertyNames -contains "device") -and $config.device.enabled -eq $true) {
        return $true
    }
    if ($hasIntegratedConfig -and $config.windows_companion.devices) {
        foreach ($property in $config.windows_companion.devices.PSObject.Properties) {
            if ($property.Value.enabled -eq $true) {
                return $true
            }
        }
    }
    return $false
}

Assert-WindowsHost
$script:CurrentIdentity = Get-CurrentIdentity
if (-not $PackageSource) {
    $PackageSource = Join-Path $PSScriptRoot "..\.."
}
$InstallRoot = [IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($InstallRoot))
$ConfigRoot = Join-Path $InstallRoot "config"
$DataRoot = Join-Path $InstallRoot "data"
$LogRoot = Join-Path $InstallRoot "logs"
$ReleaseRoot = Join-Path $InstallRoot "releases"
$BinRoot = Join-Path $InstallRoot "bin"
$StatePath = Join-Path $InstallRoot "deployment-state.json"
$InstalledConfigPath = Join-Path $ConfigRoot "windows-companion.json"
$InstalledRunnerPath = Join-Path $BinRoot "Run-WindowsCompanion.ps1"

if ($Rollback) {
    $state = Read-State -LiteralPath $StatePath
    if (-not $state -or -not $state.previous_release) {
        throw "No previous companion release is available for rollback."
    }
    if (-not $PSBoundParameters.ContainsKey("TaskName") -and $state.task_name) {
        $TaskName = [string]$state.task_name
        if ($TaskName -notmatch "^[A-Za-z0-9 ._-]{1,128}$") {
            throw "Deployment state contains an invalid Scheduled Task name."
        }
    }
    $previousPython = Join-Path ([string]$state.previous_release) "venv\Scripts\python.exe"
    if (-not (Test-Path -LiteralPath $previousPython -PathType Leaf)) {
        throw "The previous release is incomplete and cannot be selected."
    }
    if (-not $state.config_path -or
        -not (Test-Path -LiteralPath ([string]$state.config_path) -PathType Leaf)) {
        throw "The installed companion configuration is missing; rollback cannot be validated."
    }
    & $previousPython -m microsoft_agent.windows_companion_service `
        --config ([string]$state.config_path) --validate-config *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "The previous companion release rejected the installed configuration."
    }
    if ($PSCmdlet.ShouldProcess($TaskName, "Roll back to $($state.previous_release)")) {
        $oldCurrent = $state.current_release
        $state.current_release = $state.previous_release
        $state.previous_release = $oldCurrent
        if (@($state.PSObject.Properties.Name) -contains "updated_at") {
            $state.updated_at = [DateTimeOffset]::UtcNow.ToString("o")
        }
        else {
            $state | Add-Member -NotePropertyName "updated_at" `
                -NotePropertyValue ([DateTimeOffset]::UtcNow.ToString("o"))
        }
        Write-State -State $state -LiteralPath $StatePath
        Register-CompanionTask -RunnerPath $InstalledRunnerPath -StatePath $StatePath
        if (-not $NoStart) {
            Start-ScheduledTask -TaskName $TaskName
        }
    }
    [PSCustomObject]@{
        TaskName       = $TaskName
        CurrentRelease = $state.current_release
        RolledBack     = $true
    }
    return
}

$PackageSource = Resolve-ExistingPath -LiteralPath $PackageSource -Description "Package source"
if (-not (Test-Path -LiteralPath (Join-Path $PackageSource "pyproject.toml") -PathType Leaf)) {
    throw "PackageSource must be the microsoft-agent project directory containing pyproject.toml."
}
$pythonVersion = Test-PythonPrerequisite

if ($ConfigPath) {
    $ConfigPath = Resolve-ExistingPath -LiteralPath $ConfigPath -Description "Companion configuration"
}
elseif (-not (Test-Path -LiteralPath $InstalledConfigPath -PathType Leaf)) {
    $ConfigPath = Resolve-ExistingPath `
        -LiteralPath (Join-Path $PSScriptRoot "windows-companion.sample.json") `
        -Description "Sample companion configuration"
}

$releaseId = "{0}-{1}" -f (Get-Date -Format "yyyyMMddHHmmss"), $PID
$ReleasePath = Join-Path $ReleaseRoot $releaseId
$VenvPath = Join-Path $ReleasePath "venv"
$VenvPython = Join-Path $VenvPath "Scripts\python.exe"

if ($WhatIfPreference) {
    $null = $PSCmdlet.ShouldProcess($ReleasePath, "Create a Python $pythonVersion virtual environment and install microsoft-agent[windows,documents,cloud,mcp]")
    $null = $PSCmdlet.ShouldProcess($InstalledConfigPath, "Install and protect the companion configuration")
    $null = $PSCmdlet.ShouldProcess($TaskName, "Register a limited, interactive-user Scheduled Task with restart policy")
    return
}

if (-not $PSCmdlet.ShouldProcess($InstallRoot, "Install the Windows companion")) {
    return
}

$directories = @($InstallRoot, $ConfigRoot, $DataRoot, $LogRoot, $ReleaseRoot, $BinRoot, $ReleasePath)
foreach ($directory in $directories) {
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    Protect-Directory -LiteralPath $directory
}

$installCompleted = $false
$taskRegistered = $false
$oldState = $null
$stagedConfigPath = $null
try {
    Invoke-SelectedPython -Arguments @("-m", "venv", $VenvPath)
    & $VenvPython -m pip install "${PackageSource}[windows,documents,cloud,mcp]"
    if ($LASTEXITCODE -ne 0) {
        throw "Installing microsoft-agent into the companion virtual environment failed."
    }
    & $VenvPython -m microsoft_agent.windows_companion_service --help *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "The installed package does not provide a valid Windows companion service entry point."
    }

    foreach ($scriptName in @(
            "Install-WindowsCompanion.ps1",
            "Run-WindowsCompanion.ps1",
            "Test-WindowsCompanionDeployment.ps1",
            "Uninstall-WindowsCompanion.ps1"
        )) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $scriptName) `
            -Destination (Join-Path $BinRoot $scriptName) -Force
        Protect-File -LiteralPath (Join-Path $BinRoot $scriptName)
    }

    $configurationToValidate = $InstalledConfigPath
    if ($ConfigPath -and ($ReplaceConfig -or -not (Test-Path -LiteralPath $InstalledConfigPath -PathType Leaf)) -and
        [IO.Path]::GetFullPath($ConfigPath) -ne [IO.Path]::GetFullPath($InstalledConfigPath)) {
        $stagedConfigPath = "$InstalledConfigPath.new"
        Copy-Item -LiteralPath $ConfigPath -Destination $stagedConfigPath -Force
        Protect-File -LiteralPath $stagedConfigPath
        $configurationToValidate = $stagedConfigPath
    }
    if (-not (Test-Path -LiteralPath $configurationToValidate -PathType Leaf)) {
        throw "No companion configuration was installed. Supply -ConfigPath."
    }
    Protect-File -LiteralPath $configurationToValidate
    $configReady = Test-ConfigurationReadiness -LiteralPath $configurationToValidate
    & $VenvPython -m microsoft_agent.windows_companion_service `
        --config $configurationToValidate --validate-config *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "The installed companion service rejected the configuration."
    }
    if ($stagedConfigPath) {
        Move-Item -LiteralPath $stagedConfigPath -Destination $InstalledConfigPath -Force
        $stagedConfigPath = $null
        Protect-File -LiteralPath $InstalledConfigPath
    }

    $oldState = Read-State -LiteralPath $StatePath
    $oldCurrent = if ($oldState) { [string]$oldState.current_release } else { $null }
    $now = [DateTimeOffset]::UtcNow.ToString("o")
    $state = [ordered]@{
        schema_version    = 1
        installed_at     = $now
        updated_at       = $now
        current_release  = $ReleasePath
        previous_release = $oldCurrent
        config_path      = $InstalledConfigPath
        data_root        = $DataRoot
        log_root         = $LogRoot
        task_name        = $TaskName
        python_module    = "microsoft_agent.windows_companion_service"
    }
    Register-CompanionTask -RunnerPath $InstalledRunnerPath -StatePath $StatePath
    $taskRegistered = $true
    Write-State -State $state -LiteralPath $StatePath

    $installCompleted = $true
    if (-not $NoStart -and $configReady) {
        Start-ScheduledTask -TaskName $TaskName
    }
    elseif (-not $configReady) {
        Write-Warning "The installed sample device is disabled. Enroll the device, update the configuration, validate it, and then start the Scheduled Task."
    }
}
finally {
    if ($stagedConfigPath -and (Test-Path -LiteralPath $stagedConfigPath)) {
        Remove-Item -LiteralPath $stagedConfigPath -Force -ErrorAction SilentlyContinue
    }
    if (-not $installCompleted -and (Test-Path -LiteralPath $ReleasePath)) {
        Remove-Item -LiteralPath $ReleasePath -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (-not $installCompleted -and $taskRegistered) {
        $oldTaskName = if ($oldState -and
            @($oldState.PSObject.Properties.Name) -contains "task_name") {
            [string]$oldState.task_name
        }
        else {
            $null
        }
        if (-not $oldState -or $oldTaskName -ne $TaskName) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false `
                -ErrorAction SilentlyContinue
        }
    }
}

[PSCustomObject]@{
    TaskName       = $TaskName
    InstallRoot    = $InstallRoot
    ConfigPath     = $InstalledConfigPath
    CurrentRelease = $ReleasePath
    PythonVersion  = $pythonVersion.ToString()
    Started        = (-not $NoStart -and $configReady)
}
