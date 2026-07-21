# Microsoft service integrations

The default tool groups cover the daily Microsoft 365 surface while leaving
side effects disabled. Use `get_microsoft_configuration` and each integration's
configuration/capability tool before authentication.

| Service | Implemented path | Representative tools |
|---|---|---|
| Outlook mail | Microsoft Graph | `list_mail_messages`, `send_mail`, drafts, folders, attachments, move/update/delete |
| Outlook calendar | Microsoft Graph | calendars, events, calendar view, invitations and updates |
| Teams | Microsoft Graph | chats, chat messages, teams, channels, replies, members, meetings, presence |
| OneDrive/SharePoint | Microsoft Graph | drives, items, lists/sites, download/upload, generated-document upload |
| Excel | Graph workbook APIs | workbooks, worksheets, tables, ranges, formatting, sorting, charts |
| OneNote/Planner/To Do/contacts | Microsoft Graph | notes, tasks/plans, contacts and related resources |
| Word | Local OOXML generation, Graph upload, paired Office.js bridge, optional Windows COM | create/fill `.docx`, upload, agent-driven current-selection/placeholders, open/export PDF |
| PowerPoint | Local OOXML generation, Graph upload, paired Office.js bridge, optional Windows COM | create/fill `.pptx`, upload, agent-driven list/add/delete slides and text boxes, open/export PDF |
| Power Automate cloud | Supported Dataverse Web API plus named OAuth triggers | list/get/activate/deactivate solution flows, run allowlisted flow |
| Power Automate Desktop | Documented Dataverse desktop-flow APIs | list/schema/run/status/outputs/cancel for configured named desktop flows |
| Intune | Microsoft Graph v1.0 | allowlisted device/app inventory, sync, lock, reboot, shutdown, Defender scan |
| Windows laptops | Authenticated outbound companion | inventory, bounded files, Office, services, notifications, clipboard, typed extension hook |

## Document generation and upload

Install the document dependencies with `pip install .[documents]` or `.[all]`.
Templates must remain under `MICROSOFT_DOCUMENT_TEMPLATE_ROOT`; file outputs
must remain under `MICROSOFT_DOCUMENT_ARTIFACT_ROOT`. Macro-enabled packages,
embedded objects, external relationships, zip bombs, symlink escapes, and
unbounded files are rejected by default.

The combined generation/upload tools create content in memory and upload to a
drive-relative OneDrive or SharePoint path. Uploads over 10 MiB use sequential,
resumable Graph sessions; preauthenticated upload URLs never receive the Graph
bearer token. Graph [documents upload-session URLs as
opaque](https://learn.microsoft.com/en-us/graph/api/resources/uploadsession?view=graph-rest-1.0),
so the connector does not hardcode a tenant or storage hostname. It rejects
unsafe URL forms and private IP literals, then uses the shared DNS-pinned
transport to reject private/reserved DNS answers and verify the connected peer
without a second resolver lookup.

## Power Platform configuration

Set a Dataverse organization root URL or place a `power_platform` object in the
integration configuration file:

```text
MICROSOFT_DATAVERSE_ENVIRONMENT_URL=https://contoso.crm.dynamics.com
MICROSOFT_POWER_PLATFORM_ALLOW_LIFECYCLE_CHANGES=false
```

The object may also select one deployment-owned Agent Utilities TLS profile with
`tls_profile` or `tls_profile_ref`. The fields are mutually exclusive and TLS
verification cannot be disabled.

Cloud-flow definitions are read and lifecycle-managed through Dataverse's
`workflows` entity. The unsupported `api.flow.microsoft.com` management API is
rejected. Running a flow uses an administrator-configured map of names to
OAuth-protected HTTP trigger URLs; callers can supply a name and payload but
never a URL. Each invocation carries an idempotency key and correlation ID.

Desktop flows use Dataverse's documented
[`workflows`, `RunDesktopFlow`, and `flowsessions` APIs](https://learn.microsoft.com/en-us/power-automate/developer/desktop-flow-public-apis).
Configure each runnable flow by name with a fixed workflow
ID and fixed connection or connection-reference name. Execution and
cancellation have separate fail-closed flags, run modes are allowlisted per
flow, inputs are capped at Dataverse's 2 MiB limit, and callback URLs are not
accepted. The returned flow-session ID can be used to read status and outputs.
This API targets the machine or machine group bound to the configured Power
Automate connection, so it works without an arbitrary local executable hook.

For flows without an OAuth-protected request trigger, add a small solution flow
that receives the approved request and calls the target flow. Keep every URL in
the protected local configuration/secret provider even when it is OAuth
protected.

## Intune behavior

Intune configuration requires at least one managed-device UUID and an explicit
action allowlist. Every mutation requires evidence bound to the device, action,
correlation ID, idempotency key, approver, reason, and a short expiration. Reboot
and shutdown additionally require destructive acknowledgement.

The service uses only Microsoft Graph v1.0. `rotateBitLockerKeys` is visible in
the capability report as unsupported because Microsoft currently documents it
only under Graph beta; this project does not silently cross that boundary.

Copy `deployment/integrations.example.json` to a filename ending in
`.local.integrations.json`, replace the disabled examples, protect it with
user-only filesystem permissions, and set `MICROSOFT_INTEGRATIONS_CONFIG_PATH`.
The optional Windows companion object uses the same mutually exclusive TLS profile
selectors. Keep those selectors and any trust-bundle reference in the external
deployment file or secret provider rather than in source control.
