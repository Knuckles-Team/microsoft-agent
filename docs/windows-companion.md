# Windows laptop companion

Microsoft Graph and Intune cannot safely expose every local laptop service.
The companion fills that gap without opening an inbound laptop port:

```text
MCP/agent -> authenticated HTTPS control plane -> queued action
                                              <- outbound laptop poll
          <- result/health                    <- local policy + execution
```

The controller and laptop both enforce a closed action schema. There is no
arbitrary shell, PowerShell, command line, process, or URL operation.

Supported typed actions include bounded inventory, file list/read/write under
logical root bindings, Word/PowerPoint/Excel open and PDF export, allowlisted
Windows service status/start/stop, text clipboard, local notification, and an
injected allowlisted Power Automate Desktop executor. Sensitive reads and every
state change require the configured confirmation policy.

The native runtime independently verifies:

- tenant, Entra device ID, certificate thumbprint, and relay identity;
- action, path, service, and desktop-flow allowlists;
- request/confirmation action binding and expiration;
- symlink, reparse-point, device-namespace, alternate-data-stream, and root
  replacement protections for file access;
- content limits, optional expected SHA-256, atomic writes, and idempotency.

Office COM, service control, clipboard, and notification adapters are optional
and imported only on Windows. Run the worker as the enrolled interactive user,
not LocalSystem, because Office COM and the clipboard live in the user session.
The deployment assets under `deployment/windows` install a per-user Scheduled
Task and do not open firewall ports.

The control plane stores bounded action/result state in SQLite and authenticates
both controller and device requests. Terminate TLS with a trusted certificate;
when a reverse proxy supplies an mTLS certificate thumbprint, accept that header
only from an explicitly configured trusted proxy. See the deployment README for
the exact worker command and enrollment sequence. The runnable relay entry point
is `microsoft-windows-control-plane`; its disabled, secret-free sample and
deployment instructions are under `deployment/control-plane`.
