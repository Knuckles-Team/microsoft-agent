# Windows companion deployment

This bundle installs the native Windows companion as a limited, per-user
Scheduled Task. The companion makes authenticated outbound HTTPS requests to a
control plane; it does not listen on a laptop port, create a firewall rule, or
provide a shell/process action.

After enrollment it polls and acknowledges the relay unattended for as long as
the user session exists. Polling carries an Entra application token and mTLS
client identity over outbound HTTPS only; action payloads never arrive through
an inbound laptop listener.

The task deliberately does **not** run as `SYSTEM`. Word/PowerPoint/Excel COM,
the clipboard, notifications, and Power Automate Desktop belong to the signed-in
desktop session and are unreliable or unsafe from a service/session-0 process.

## Prerequisites

- Windows 10 or 11 with Windows PowerShell 5.1 and Task Scheduler.
- 64-bit Python 3.11 through 3.14, available through `py.exe` or an explicit
  `-PythonCommand` path.
- The local `microsoft-agent` source tree or an internally reviewed source
  checkout. The installer runs `pip install
  <source>[windows,documents,cloud,mcp]` inside a release-specific virtual
  environment. It does not download an executable directly.
- Outbound TCP 443 access to the approved control-plane origin and Microsoft
  identity endpoints. No inbound port is required.
- Desktop Microsoft Office only for Office actions. The rest of the companion
  can run without Office installed if those actions are not allowlisted.

Keep organizational PowerShell signing and execution policies in force. These
scripts never set `ExecutionPolicy`, disable Defender, weaken TLS, or change the
firewall.

## Entra and certificate enrollment

Authentication can be added after the software is installed. Leave the sample
device disabled until every identity value has been provisioned.

1. Register a single-tenant Entra client application for native companions,
   separate from the control-plane API application. Give the client only the
   control-plane `WindowsCompanion.Device` application role required to poll
   and acknowledge actions. Its token resource is
   `api://<control-plane-api-client-id>`; it does not need Microsoft Graph
   permissions merely to run the companion.
2. Issue a separate client certificate to each laptop. Add the public
   credential to the Entra application (or use a per-device workload identity)
   and enroll the public mTLS certificate/CA with the control plane.
3. Record the tenant ID, client ID, Entra device ID, logical device ID, token
   audience, certificate thumbprint, and certificate/private-key **paths** in a
   copy of `windows-companion.sample.json`. Never add private-key contents,
   passwords, access tokens, or client secrets to source control.
4. Place the PEM private key and certificate in a local staging directory
   readable only by the installing user, then configure their final protected
   paths as required by the service configuration. A private key without a PEM
   password is acceptable only inside the installer-created ACL boundary and
   under your endpoint-management policy. For an encrypted PEM, inject
   `MICROSOFT_COMPANION_CERTIFICATE_PASSWORD` into the task process with your
   approved Windows credential broker. The supplied scripts intentionally do
   not collect or persist that password.
5. Configure the control plane to require both the expected Entra token claims
   and the enrolled client certificate, and to bind them to the configured
   `device_id`. Enable the device only after that binding is complete.

The no-secret helper at
`deployment/entra/New-WindowsControlPlaneEntraResources.ps1` can create the
separate v2 API and one certificate-backed worker application, declare the
controller/worker permissions, and emit the exact client/object IDs needed by
both JSON files. It does not grant admin consent without its explicit
`-GrantAdminConsent` switch.

Certificate rotation should overlap old and new public credentials briefly,
update the protected config path/thumbprint, validate, then remove the old
credential. Do not reuse a private key across laptops.

## Install and dry-run

From a normal (non-elevated) PowerShell window in this directory:

```powershell
.\Install-WindowsCompanion.ps1 -ConfigPath C:\SecureStaging\windows-companion.json -WhatIf
.\Install-WindowsCompanion.ps1 -ConfigPath C:\SecureStaging\windows-companion.json
```

`-WhatIf` validates Windows, Python, source, and path prerequisites, then shows
the release/config/task changes without writing them. Use `-NoStart` during
staged enrollment. Use `-ReplaceConfig` when an existing installed config must
be replaced explicitly.

The default installation root is:

```text
%LOCALAPPDATA%\MicrosoftAgent\Companion
```

Its config, data, logs, versioned environments, scripts, and state file have
ACL inheritance disabled and grant access only to the current user, local
`SYSTEM`, and local Administrators. The state contains paths and release
metadata, not credentials. The
Scheduled Task starts at that user's logon, retries failures five times, and
has a 15-minute watchdog trigger while the user is logged on. Multiple
instances are prohibited.

If no `-ConfigPath` is supplied on a first install, the installer copies the
sample, registers the task, and does not start it because the only sample device
is disabled.

The protected config stores identifiers, allowlists, and credential **paths**.
It may contain a certificate thumbprint but never the private-key bytes,
private-key password, token, or client secret. Certificate/key files referenced
by an enabled config must be regular non-symlink files; the service validates
the certificate/key pair and thumbprint before polling. Keep those files inside
the same current-user/SYSTEM/Administrators ACL boundary unless an enterprise
credential store applies an equally restrictive policy.

## Validate and operate

Run the read-only self-check after enrollment or an update:

```powershell
.\Test-WindowsCompanionDeployment.ps1
.\Test-WindowsCompanionDeployment.ps1 -AsJson
```

It validates protected ACLs, deployment state, the active Python version,
module loading, configuration validation when exposed by the installed
service, the current-user/limited Scheduled Task principal, logon/retry
settings, and current task health. It does not acquire a token or call the
network.

Normal task operations use standard commands:

```powershell
Start-ScheduledTask -TaskName "Microsoft Agent Windows Companion"
Stop-ScheduledTask -TaskName "Microsoft Agent Windows Companion"
Get-ScheduledTask -TaskName "Microsoft Agent Windows Companion"
Get-ScheduledTaskInfo -TaskName "Microsoft Agent Windows Companion"
```

The launcher keeps the newest 20 stdout/stderr log files under the protected
`logs` directory. Logs should contain lifecycle and safe error information;
they must not contain tokens, certificate passwords, private keys, file
contents, or clipboard values. Control-plane health is authoritative for
outbound connectivity; Task Scheduler state and these logs diagnose failures
before the companion can report health.

## Local policy and adapter boundaries

The device and controller both enforce the closed action model and allowlists.
Do not enable an action until its matching file root, service name, or desktop
flow is explicitly approved. File roots use logical Windows paths mapped to
local protected paths, and the runtime rejects symlinks, junctions, and reparse
points.

Power Automate Desktop execution is an injected adapter boundary. The runtime
does not fall back to `powershell.exe`, `cmd.exe`, arbitrary processes, URIs, or
unvalidated PAD flow names. A trusted deployment-specific
`DesktopFlowExecutor` must be installed/configured before
`power_automate_desktop.run` can be allowlisted. Startup fails closed when an
enabled action has no corresponding native adapter.

Office COM and clipboard actions require the user to be logged on. Office may
display prompts for protected, damaged, or policy-governed documents; endpoint
automation policy should decide whether those actions are allowed.

## Updates and rollback

Re-run the installer from the reviewed updated source. Each install builds a
new release directory and changes the small protected state file only after
package/module validation succeeds. The previous environment remains available:

```powershell
.\Install-WindowsCompanion.ps1 -ConfigPath C:\SecureStaging\windows-companion.json -NoStart
.\Test-WindowsCompanionDeployment.ps1
Start-ScheduledTask -TaskName "Microsoft Agent Windows Companion"

# Restore the immediately previous release without reinstalling packages.
.\Install-WindowsCompanion.ps1 -Rollback
```

Rollback swaps current and previous release pointers, so the same command can
switch back after investigation. Old release directories beyond those pointers
may be removed during a maintenance window after validation; do not delete the
active or rollback release.

## Uninstall

```powershell
.\Uninstall-WindowsCompanion.ps1 -WhatIf
.\Uninstall-WindowsCompanion.ps1
```

The default removes the task, virtual environments, runner, and deployment
state but retains protected config/data/logs. To erase the entire companion
directory after backup and certificate de-registration:

```powershell
.\Uninstall-WindowsCompanion.ps1 -PurgeData
```

Uninstallation does not delete Entra credentials or control-plane enrollment.
Revoke those separately, then securely destroy any staged or retained private
key material according to your device-management policy.
