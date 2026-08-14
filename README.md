# Microsoft Agent

Microsoft Agent provides a governed Microsoft Graph MCP server and an optional A2A
agent. It uses Agent Utilities for MCP, identity propagation, intent delegation,
configuration, telemetry, and graph integration.

Current package version: **2.1.0**

## Capabilities

- Outlook mail, calendars, contacts, tasks, and attachments
- Teams chats, channels, meetings, and presence
- OneDrive, SharePoint, Excel, OneNote, and Universal Print
- Entra users, groups, applications, policy, audit, security, and reports
- Optional Word and PowerPoint generation and an exact-origin Office.js bridge
- Optional Power Platform, Intune, and outbound Windows companion integrations
- Governed ingestion into Epistemic Graph with provider-owned ontology, mapping,
  schema fingerprint, provenance, and quarantine contracts
- One consolidated `microsoft-agent-operations` skill

The current implementation has one modular Graph client in
`microsoft_agent/api_client.py`, one authentication authority in
`microsoft_agent/auth.py`, and one MCP entry point in
`microsoft_agent/mcp_server.py`. No raw-token fallback, plaintext cache, bundled
database, bundled endpoint profile, or duplicate API wrapper is shipped.

## Install

```bash
python -m pip install "microsoft-agent[mcp]"
```

Useful extras:

```bash
python -m pip install "microsoft-agent[agent]"        # A2A agent runtime
python -m pip install "microsoft-agent[documents]"    # Word and PowerPoint
python -m pip install "microsoft-agent[cloud]"        # managed/workload identity
python -m pip install "microsoft-agent[windows]"      # Windows companion
python -m pip install "microsoft-agent[control-plane]"
python -m pip install "microsoft-agent[all]"
```

Agent Utilities depends on `epistemic-graph[full]`; the full and numeric engine
features therefore remain available when the A2A/GraphOS runtime is installed.

## Authentication

Microsoft identity values are supplied by deployment configuration. There is no
baked-in client identifier.

Supported modes:

- `delegated`: broker or browser login; device code is opt-in
- `application`: certificate or referenced client secret
- `on_behalf_of`: verified incoming user delegation
- `external_token`: verified request token for an explicitly allowed audience
- `managed_identity`: Azure managed identity
- `workload_identity`: federated workload token file

Token caches use OS-backed secure storage only. If secure storage is unavailable,
tokens remain in memory and are not written to plaintext files.

<!-- ENV-VARS-TABLE:START -->

#### Package environment variables

| Variable | Example | Description |
|----------|---------|-------------|
| `MICROSOFT_TENANT_ID` | — | Microsoft identity (values intentionally blank; supply them after app enrollment) |
| `MICROSOFT_CLIENT_ID` | — |  |
| `MICROSOFT_GRAPH_TLS_PROFILE` | — |  |
| `MICROSOFT_GRAPH_TLS_PROFILE_REF` | — |  |
| `MICROSOFT_AUTH_MODE` | `delegated` | delegated \| application \| on_behalf_of \| external_token \| managed_identity \| workload_identity |
| `MICROSOFT_LOGIN_METHOD` | `auto` |  |
| `MICROSOFT_ENABLE_BROKER` | `true` |  |
| `MICROSOFT_ALLOW_DEVICE_CODE` | `false` |  |
| `MICROSOFT_REQUIRE_SECURE_CACHE` | `true` |  |
| `MICROSOFT_PERMISSION_PROFILES` | `productivity,collaboration` | Named least-privilege bundles. Add device_read/device_admin only for Intune. |
| `MICROSOFT_GRAPH_SCOPES` | — |  |
| `MICROSOFT_ENABLED_TOOL_GROUPS` | `misc,auth,meta,mail,files,calendar,notes,tasks,contacts,user,chat,teams,sites,search,groups,communications,documents,power_platform,windows,intune` | Tools are registered, but side effects remain disabled until explicitly enabled. |
| `MICROSOFT_ALLOW_WRITES` | `false` |  |
| `MICROSOFT_ALLOW_DESTRUCTIVE` | `false` |  |
| `MICROSOFT_INGESTION_PSEUDONYMIZATION_KEY_REF` | — | Required only for graph projection/ingestion. Resolve from AgentConfig or a secret store; use at least 32 bytes and never reuse an identity credential. |
| `MICROSOFT_INTEGRATIONS_CONFIG_PATH` | — | Complex Power Platform, Intune, and Windows companion allowlists. |
| `MICROSOFT_DOCUMENT_ARTIFACT_ROOT` | — |  |
| `MICROSOFT_DOCUMENT_TEMPLATE_ROOT` | — |  |
| `MICROSOFT_DATAVERSE_ENVIRONMENT_URL` | — |  |
| `MICROSOFT_DATAVERSE_AUDIENCE` | — |  |
| `MICROSOFT_POWER_AUTOMATE_NAMED_FLOWS_JSON` | `{}` |  |
| `MICROSOFT_POWER_PLATFORM_ALLOW_LIFECYCLE_CHANGES` | `false` |  |
| `MICROSOFT_OFFICE_ADDIN_ORIGINS` | — |  |
| `MICROSOFT_CLIENT_CERTIFICATE_PATH` | — | Confidential modes: prefer a certificate, managed identity, or workload identity. |
| `MICROSOFT_CLIENT_CERTIFICATE_THUMBPRINT` | — |  |
| `MICROSOFT_CLIENT_CERTIFICATE_PASSWORD_REF` | — |  |
| `MICROSOFT_CLIENT_SECRET_REF` | — |  |
| `MICROSOFT_MANAGED_IDENTITY_CLIENT_ID` | — |  |
| `MICROSOFT_WORKLOAD_IDENTITY_TOKEN_FILE` | secret-injected |  |
| `TRANSPORT` | `stdio` | Local MCP defaults. Require an authenticated transport before non-loopback use. |
| `HOST` | `127.0.0.1` |  |
| `PORT` | `8000` |  |
| `AUTH_TYPE` | `oidc` |  |

#### Inherited agent-utilities variables (apply to every connector)

| Variable | Example | Description |
|----------|---------|-------------|
| `MCP_TOOL_MODE` | `intent` | Tool surface: `intent` \| `condensed` \| `verbose` \| `both` |
| `MCP_ENABLED_TOOLS` | — | Comma-separated tool allow-list |
| `MCP_DISABLED_TOOLS` | — | Comma-separated tool deny-list |
| `MCP_ENABLED_TAGS` | — | Comma-separated tag allow-list |
| `MCP_DISABLED_TAGS` | — | Comma-separated tag deny-list |
| `EUNOMIA_TYPE` | `none` | Authorization mode: `none` \| `embedded` \| `remote` |
| `EUNOMIA_POLICY_FILE` | `mcp_policies.json` | Embedded Eunomia policy file |
| `EUNOMIA_REMOTE_URL` | — | Remote Eunomia authorization server URL |
| `ENABLE_OTEL` | `False` | Enable OpenTelemetry export |
| `OTEL_EXPORTER_OTLP_ENDPOINT` | — | OTLP collector endpoint |
| `MCP_CLIENT_AUTH` | — | Outbound MCP child auth: `oidc-client-credentials` \| `basic` \| `none` |
| `OIDC_CLIENT_ID` | — | OIDC client id (service-account auth) |
| `OIDC_CLIENT_SECRET_REF` | `secret://identity/oidc-client-secret` | Runtime secret reference for the OIDC service account |
| `MCP_BASIC_AUTH_USERNAME` | — | HTTP Basic username (`MCP_CLIENT_AUTH=basic`) |
| `MCP_BASIC_AUTH_PASSWORD_REF` | `secret://identity/mcp-basic-password` | Runtime secret reference for HTTP Basic auth (`MCP_CLIENT_AUTH=basic`) |
| `DEBUG` | `False` | Verbose logging |
| `PYTHONUNBUFFERED` | `1` | Unbuffered stdout (recommended in containers) |
| `MCP_URL` | `http://localhost:8000/mcp` | URL of the MCP server the agent connects to |
| `PROVIDER` | `openai` | LLM provider for the agent |
| `MODEL_ID` | `gpt-4o` | Model id for the agent |
| `ENABLE_WEB_UI` | `True` | Serve the AG-UI web interface |

_33 package + 21 inherited variable(s). Auto-generated from `.env.example` + the shared agent-utilities set — do not edit._
<!-- ENV-VARS-TABLE:END -->

## Run the MCP server

Local stdio:

```bash
microsoft-mcp --transport stdio
```

Authenticated HTTP deployment:

```bash
microsoft-mcp --transport streamable-http --host 127.0.0.1 --port 8000
```

Do not expose an unauthenticated listener outside loopback. Microsoft Graph identity
and MCP caller identity are separate boundaries and both must be configured.

<!-- MCP-TOOLS-TABLE:START -->

#### Condensed action-routed tools (`MCP_TOOL_MODE=condensed`)

| MCP Tool | Toggle Env Var | Description |
|----------|----------------|-------------|
| `activate_power_automate_flow` | `POWER_PLATFORMTOOL` | Activate a solution-aware cloud flow through Dataverse. |
| `add_powerpoint_slide_in_office` | `DOCUMENTSTOOL` | Append a slide in a paired, open PowerPoint task pane. |
| `cancel_power_automate_desktop_flow_run` | `POWER_PLATFORMTOOL` | Cancel a queued or running desktop flow through Dataverse when cancellation is explicitly enabled. |
| `create_office_pairing` | `DOCUMENTSTOOL` | Create a five-minute, one-time pairing secret for a visible Word or PowerPoint task pane. Give the secret only to the intended user. |
| `deactivate_power_automate_flow` | `POWER_PLATFORMTOOL` | Deactivate a solution-aware cloud flow through Dataverse. |
| `delete_powerpoint_slide_in_office` | `DOCUMENTSTOOL` | Delete one numbered slide in a paired, open PowerPoint task pane. |
| `generate_and_upload_powerpoint_presentation` | `DOCUMENTTOOL` | Generate a PowerPoint .pptx presentation in memory and upload it to a OneDrive or SharePoint document-library drive path. |
| `generate_and_upload_word_document` | `DOCUMENTTOOL` | Generate a Word .docx document in memory and upload it to a OneDrive or SharePoint document-library drive path. |
| `generate_powerpoint_presentation` | `DOCUMENTTOOL` | Generate a non-macro PowerPoint .pptx presentation from validated slides, metadata, or a template confined to the configured root. |
| `generate_word_document` | `DOCUMENTTOOL` | Generate a non-macro Word .docx file from validated paragraphs, tables, metadata, or a template confined to the configured root. |
| `get_document_capabilities` | `DOCUMENTTOOL` | Report whether local Word and PowerPoint OOXML generation backends are installed. This tool does not require Microsoft authentication. |
| `get_intune_configuration` | `INTUNETOOL` | Return sanitized Intune readiness, device count, and stable v1.0 remote-action capabilities. |
| `get_intune_managed_device` | `INTUNETOOL` | Get inventory and compliance data for one allowlisted Intune device. |
| `get_office_command_result` | `DOCUMENTSTOOL` | Get the state or retained typed result of an Office bridge command. |
| `get_power_automate_desktop_flow_outputs` | `POWER_PLATFORMTOOL` | Read outputs for a completed Power Automate desktop-flow run. |
| `get_power_automate_desktop_flow_run` | `POWER_PLATFORMTOOL` | Get status and timestamps for a Power Automate desktop-flow run. |
| `get_power_automate_desktop_flow_schema` | `POWER_PLATFORMTOOL` | Get the input or output schema for one configured desktop flow. The caller selects an allowlisted name, never a workflow URL. |
| `get_power_automate_flow` | `POWER_PLATFORMTOOL` | Get one solution-aware cloud-flow definition from Dataverse. |
| `get_power_platform_configuration` | `POWER_PLATFORMTOOL` | Return sanitized Power Platform readiness and named-flow status. |
| `get_windows_action_result` | `WINDOWS_COMPANIONTOOL` | Read current or final state for a submitted Windows action. |
| `get_windows_companion_configuration` | `WINDOWS_COMPANIONTOOL` | Return sanitized Windows companion device and action readiness. |
| `get_windows_device_health` | `WINDOWS_COMPANIONTOOL` | Read authenticated outbound-relay health for an allowlisted laptop. |
| `get_word_selection_from_office` | `DOCUMENTSTOOL` | Read the current selection from a paired, open Word task pane. |
| `insert_powerpoint_text_box_in_office` | `DOCUMENTSTOOL` | Insert a positioned text box in a paired, open PowerPoint task pane. |
| `list_intune_detected_apps` | `INTUNETOOL` | List tenant-wide Intune detected applications only when that sensitive inventory capability is explicitly enabled. |
| `list_intune_managed_devices` | `INTUNETOOL` | List only managed-device IDs explicitly allowlisted for this agent. |
| `list_microsoft_ingestion_projection` | `KGTOOL` | Return only keyed opaque nodes and structural relationships. |
| `list_office_sessions` | `DOCUMENTSTOOL` | List paired, currently connected Office task panes without secrets. |
| `list_power_automate_desktop_flows` | `POWER_PLATFORMTOOL` | List published Power Automate desktop flows through the documented Dataverse workflow API; draft inclusion is explicit. |
| `list_power_automate_flows` | `POWER_PLATFORMTOOL` | List solution-aware cloud flows through the Dataverse Web API. |
| `list_powerpoint_slides_in_office` | `DOCUMENTSTOOL` | List slides in a paired, open PowerPoint task pane. |
| `login_power_platform` | `POWER_PLATFORMTOOL` | Explicitly acquire delegated tokens for the configured Dataverse and named-flow resource audiences. May open the broker/browser. |
| `login_windows_companion` | `WINDOWS_COMPANIONTOOL` | Explicitly acquire a delegated token for the configured Windows companion control-plane audience. May open the broker/browser. |
| `microsoft_admin` | `ADMINTOOL` | Manage microsoft admin operations. |
| `microsoft_agreements` | `AGREEMENTSTOOL` | Manage microsoft agreements operations. |
| `microsoft_applications` | `APPLICATIONSTOOL` | Manage microsoft applications operations. |
| `microsoft_audit` | `AUDITTOOL` | Manage microsoft audit operations. |
| `microsoft_auth` | `AUTHTOOL` | Manage microsoft auth operations. |
| `microsoft_calendar` | `CALENDARTOOL` | Manage microsoft calendar operations. |
| `microsoft_chat` | `CHATTOOL` | Manage microsoft chat operations. |
| `microsoft_communications` | `COMMUNICATIONSTOOL` | Manage microsoft communications operations. |
| `microsoft_connections` | `CONNECTIONSTOOL` | Manage microsoft connections operations. |
| `microsoft_contacts` | `CONTACTSTOOL` | Manage microsoft contacts operations. |
| `microsoft_devices` | `DEVICESTOOL` | Manage microsoft devices operations. |
| `microsoft_directory` | `DIRECTORYTOOL` | Manage microsoft directory operations. |
| `microsoft_domains` | `DOMAINSTOOL` | Manage microsoft domains operations. |
| `microsoft_education` | `EDUCATIONTOOL` | Manage microsoft education operations. |
| `microsoft_employee_experience` | `EMPLOYEE_EXPERIENCETOOL` | Manage microsoft employee experience operations. |
| `microsoft_files` | `FILESTOOL` | Manage microsoft files operations. |
| `microsoft_groups` | `GROUPSTOOL` | Manage microsoft groups operations. |
| `microsoft_identity` | `IDENTITYTOOL` | Manage microsoft identity operations. |
| `microsoft_mail` | `MAILTOOL` | Manage microsoft mail operations. |
| `microsoft_meta` | `METATOOL` | Manage microsoft meta operations. |
| `microsoft_notes` | `NOTESTOOL` | Manage microsoft notes operations. |
| `microsoft_organization` | `ORGANIZATIONTOOL` | Manage microsoft organization operations. |
| `microsoft_places` | `PLACESTOOL` | Manage microsoft places operations. |
| `microsoft_policies` | `POLICIESTOOL` | Manage microsoft policies operations. |
| `microsoft_print` | `PRINTTOOL` | Manage microsoft print operations. |
| `microsoft_privacy` | `PRIVACYTOOL` | Manage microsoft privacy operations. |
| `microsoft_reports` | `REPORTSTOOL` | Manage microsoft reports operations. |
| `microsoft_search` | `SEARCHTOOL` | Manage microsoft search operations. |
| `microsoft_security` | `SECURITYTOOL` | Manage microsoft security operations. |
| `microsoft_sites` | `SITESTOOL` | Manage microsoft sites operations. |
| `microsoft_solutions` | `SOLUTIONSTOOL` | Manage microsoft solutions operations. |
| `microsoft_storage` | `STORAGETOOL` | Manage microsoft storage operations. |
| `microsoft_subscriptions` | `SUBSCRIPTIONSTOOL` | Manage microsoft subscriptions operations. |
| `microsoft_tasks` | `TASKSTOOL` | Manage microsoft tasks operations. |
| `microsoft_teams` | `TEAMSTOOL` | Manage microsoft teams operations. |
| `microsoft_user` | `USERTOOL` | Manage microsoft user operations. |
| `reboot_intune_device` | `INTUNETOOL` | Immediately reboot an allowlisted Intune device after destructive action acknowledgement and time-bound confirmation. |
| `remote_lock_intune_device` | `INTUNETOOL` | Remotely lock an allowlisted Intune device after time-bound confirmation bound to that device and action. |
| `replace_word_placeholders_in_office` | `DOCUMENTSTOOL` | Replace a literal placeholder throughout a paired, open Word document. |
| `run_power_automate_desktop_flow` | `POWER_PLATFORMTOOL` | Run one allowlisted Power Automate desktop flow using its fixed Dataverse connection or connection reference. |
| `run_power_automate_flow` | `POWER_PLATFORMTOOL` | Invoke one named, allowlisted, OAuth-protected Power Automate HTTP trigger. The caller cannot supply a trigger URL. |
| `scan_intune_device_with_defender` | `INTUNETOOL` | Request a quick or full Microsoft Defender scan on an allowlisted Intune device with time-bound confirmation. |
| `shut_down_intune_device` | `INTUNETOOL` | Immediately shut down an allowlisted Intune device after destructive action acknowledgement and time-bound confirmation. |
| `submit_windows_action` | `WINDOWS_COMPANIONTOOL` | Submit one typed allowlisted laptop action through the authenticated outbound relay. Arbitrary shell, process, and URL execution are absent. |
| `sync_intune_device` | `INTUNETOOL` | Request an Intune sync for an allowlisted device. Time-bound confirmation evidence bound to the action and device is required. |
| `write_word_selection_in_office` | `DOCUMENTSTOOL` | Write bounded text at the selection in a paired, open Word task pane. |

#### Verbose 1:1 API-mapped tools (`MCP_TOOL_MODE=verbose` or `both`)

<details>
<summary>268 per-operation tools — one per public API method (click to expand)</summary>

| MCP Tool | Toggle Env Var | Description |
|----------|----------------|-------------|
| `microsoft_add_application_password` | `OTHERTOOL` | Add a password credential to an application. |
| `microsoft_add_group_member` | `DIRECTORYTOOL` | Add a member to a group. |
| `microsoft_add_mail_attachment` | `MAILTOOL` | Add attachment to message. |
| `microsoft_create_agreement` | `OTHERTOOL` | Create an agreement. |
| `microsoft_create_application` | `OTHERTOOL` | Create an application registration. |
| `microsoft_create_booking_appointment` | `OTHERTOOL` | Create a booking appointment. |
| `microsoft_create_calendar_event` | `CALENDARTOOL` | Create calendar event. |
| `microsoft_create_conditional_access_policy` | `ADMINTOOL` | Create a conditional access policy. |
| `microsoft_create_domain` | `ADMINTOOL` | Add a domain to the tenant. |
| `microsoft_create_draft_email` | `MAILTOOL` | Create draft email. |
| `microsoft_create_excel_chart` | `APPSTOOL` | Create a chart in an Excel worksheet. |
| `microsoft_create_external_connection` | `OTHERTOOL` | Create an external connection. |
| `microsoft_create_file_storage_container` | `DRIVETOOL` | Create a file storage container. |
| `microsoft_create_group` | `DIRECTORYTOOL` | Create a new group. |
| `microsoft_create_invitation` | `OTHERTOOL` | Create an invitation for a guest user. |
| `microsoft_create_onenote_page` | `APPSTOOL` | Create a OneNote page from the raw HTML body required by Graph. |
| `microsoft_create_online_meeting` | `CALENDARTOOL` | Create a new online meeting. |
| `microsoft_create_outlook_contact` | `CALENDARTOOL` | Create Outlook contact. |
| `microsoft_create_planner_task` | `APPSTOOL` | Create Planner task. |
| `microsoft_create_print_document_upload_session` | `OTHERTOOL` | Create the documented preauthenticated Universal Print upload session. |
| `microsoft_create_print_job` | `OTHERTOOL` | Create a print job. |
| `microsoft_create_role_assignment` | `ADMINTOOL` | Create a role assignment. |
| `microsoft_create_service_principal` | `OTHERTOOL` | Create a service principal. |
| `microsoft_create_specific_calendar_event` | `CALENDARTOOL` | Create specific calendar event. |
| `microsoft_create_subject_rights_request` | `OTHERTOOL` | Create a subject rights request. |
| `microsoft_create_subscription` | `OTHERTOOL` | Create a subscription for change notifications. |
| `microsoft_create_todo_task` | `APPSTOOL` | Create Todo task. |
| `microsoft_delete_agreement` | `OTHERTOOL` | Delete an agreement. |
| `microsoft_delete_application` | `OTHERTOOL` | Delete an application. |
| `microsoft_delete_calendar_event` | `CALENDARTOOL` | Delete calendar event. |
| `microsoft_delete_conditional_access_policy` | `ADMINTOOL` | Delete a conditional access policy. |
| `microsoft_delete_device` | `OTHERTOOL` | Delete a device. |
| `microsoft_delete_domain` | `ADMINTOOL` | Delete a domain. |
| `microsoft_delete_external_connection` | `OTHERTOOL` | Delete an external connection. |
| `microsoft_delete_group` | `DIRECTORYTOOL` | Delete a group. |
| `microsoft_delete_mail_attachment` | `MAILTOOL` | Delete attachment. |
| `microsoft_delete_mail_message` | `MAILTOOL` | Delete a message. |
| `microsoft_delete_onedrive_file` | `DRIVETOOL` | Delete file. |
| `microsoft_delete_online_meeting` | `CALENDARTOOL` | Delete an online meeting. |
| `microsoft_delete_outlook_contact` | `CALENDARTOOL` | Delete Outlook contact. |
| `microsoft_delete_service_principal` | `OTHERTOOL` | Delete a service principal. |
| `microsoft_delete_specific_calendar_event` | `CALENDARTOOL` | Delete specific calendar event. |
| `microsoft_delete_subscription` | `OTHERTOOL` | Delete a subscription. |
| `microsoft_delete_todo_task` | `APPSTOOL` | Delete Todo task. |
| `microsoft_dismiss_risky_user` | `DIRECTORYTOOL` | Dismiss a risky user. |
| `microsoft_download_onedrive_file_content` | `DRIVETOOL` | Download file content. |
| `microsoft_find_meeting_times` | `CALENDARTOOL` | Find meeting times. |
| `microsoft_format_excel_range` | `APPSTOOL` | Update formatting for an addressed Excel worksheet range. |
| `microsoft_get_access_review` | `ADMINTOOL` | Get a specific access review definition. |
| `microsoft_get_admin_consent_policy` | `ADMINTOOL` | Get the admin consent request policy. |
| `microsoft_get_admin_sharepoint` | `DRIVETOOL` | Get SharePoint admin settings. |
| `microsoft_get_agreement` | `OTHERTOOL` | Get a specific agreement. |
| `microsoft_get_application` | `OTHERTOOL` | Get a specific application. |
| `microsoft_get_authorization_policy` | `ADMINTOOL` | Get the authorization policy. |
| `microsoft_get_booking_business` | `OTHERTOOL` | Get a specific booking business. |
| `microsoft_get_calendar_event` | `CALENDARTOOL` | Get calendar event. |
| `microsoft_get_calendar_view` | `CALENDARTOOL` | Get calendar view. |
| `microsoft_get_call_record` | `OTHERTOOL` | Get a specific call record. |
| `microsoft_get_channel_message` | `MAILTOOL` | Get channel message. |
| `microsoft_get_chat` | `DIRECTORYTOOL` | Get chat. |
| `microsoft_get_chat_message` | `MAILTOOL` | Get chat message. |
| `microsoft_get_conditional_access_policy` | `ADMINTOOL` | Get a specific conditional access policy. |
| `microsoft_get_delegated_admin_relationship` | `OTHERTOOL` | Get a specific delegated admin relationship. |
| `microsoft_get_device` | `OTHERTOOL` | Get a specific device. |
| `microsoft_get_directory_audit` | `ADMINTOOL` | Get a specific directory audit entry. |
| `microsoft_get_directory_object` | `OTHERTOOL` | Get a specific directory object. |
| `microsoft_get_directory_role` | `ADMINTOOL` | Get a specific directory role. |
| `microsoft_get_domain` | `ADMINTOOL` | Get domain details. |
| `microsoft_get_drive_root_item` | `DRIVETOOL` | Get drive root item. |
| `microsoft_get_education_class` | `OTHERTOOL` | Get a specific education class. |
| `microsoft_get_education_school` | `OTHERTOOL` | Get a specific education school. |
| `microsoft_get_email_activity_report` | `MAILTOOL` | Get email activity user detail report. |
| `microsoft_get_excel_range` | `APPSTOOL` | Get an addressed range from an Excel worksheet. |
| `microsoft_get_excel_table` | `APPSTOOL` | Get Excel table. |
| `microsoft_get_excel_workbook` | `APPSTOOL` | Get Excel workbook. |
| `microsoft_get_excel_worksheet` | `APPSTOOL` | Get Excel worksheet. |
| `microsoft_get_external_connection` | `OTHERTOOL` | Get a specific external connection. |
| `microsoft_get_file_storage_container` | `DRIVETOOL` | Get a specific file storage container. |
| `microsoft_get_group` | `DIRECTORYTOOL` | Get a specific group. |
| `microsoft_get_learning_provider` | `OTHERTOOL` | Get a specific learning provider. |
| `microsoft_get_mail_attachment` | `MAILTOOL` | Get attachment. |
| `microsoft_get_mail_message` | `MAILTOOL` | Get a specific message. |
| `microsoft_get_mailbox_usage_report` | `MAILTOOL` | Get mailbox usage detail report. |
| `microsoft_get_managed_device` | `OTHERTOOL` | Get a specific managed device. |
| `microsoft_get_me` | `OTHERTOOL` | Get the current user. |
| `microsoft_get_my_presence` | `DIRECTORYTOOL` | Get current user's presence. |
| `microsoft_get_office365_active_users` | `DIRECTORYTOOL` | Get Office 365 active user detail report. |
| `microsoft_get_onedrive_usage_report` | `DRIVETOOL` | Get OneDrive usage account detail report. |
| `microsoft_get_onenote_page_content` | `APPSTOOL` | Get Onenote page content. |
| `microsoft_get_online_meeting` | `CALENDARTOOL` | Get a specific online meeting. |
| `microsoft_get_org_branding` | `OTHERTOOL` | Get organization branding. |
| `microsoft_get_organization` | `ADMINTOOL` | Get organization by ID. |
| `microsoft_get_outlook_contact` | `CALENDARTOOL` | Get Outlook contact. |
| `microsoft_get_place` | `OTHERTOOL` | Get a specific place. |
| `microsoft_get_planner_plan` | `APPSTOOL` | Get Planner plan. |
| `microsoft_get_planner_task` | `APPSTOOL` | Get Planner task. |
| `microsoft_get_presence` | `DIRECTORYTOOL` | Get presence for a specific user. |
| `microsoft_get_printer` | `OTHERTOOL` | Get a specific printer. |
| `microsoft_get_risk_detection` | `OTHERTOOL` | Get a specific risk detection. |
| `microsoft_get_risky_user` | `DIRECTORYTOOL` | Get a specific risky user. |
| `microsoft_get_role_assignment` | `ADMINTOOL` | Get a specific role assignment. |
| `microsoft_get_role_definition` | `ADMINTOOL` | Get a specific role definition. |
| `microsoft_get_security_alert` | `ADMINTOOL` | Get a specific security alert. |
| `microsoft_get_security_incident` | `ADMINTOOL` | Get a specific security incident. |
| `microsoft_get_sensitivity_label` | `OTHERTOOL` | Get a specific sensitivity label. |
| `microsoft_get_service_health` | `ADMINTOOL` | Get service health for a specific service. |
| `microsoft_get_service_health_issue` | `ADMINTOOL` | Get a specific service health issue. |
| `microsoft_get_service_principal` | `OTHERTOOL` | Get a specific service principal. |
| `microsoft_get_service_update_message` | `MAILTOOL` | Get a specific service update message. |
| `microsoft_get_shared_mailbox_message` | `MAILTOOL` | Get a message from a shared mailbox. |
| `microsoft_get_sharepoint_activity_report` | `DRIVETOOL` | Get SharePoint activity user detail report. |
| `microsoft_get_sharepoint_site_by_path` | `DRIVETOOL` | Get SharePoint site by path. |
| `microsoft_get_sharepoint_site_list_item` | `DRIVETOOL` | Get an item in a SharePoint site list. |
| `microsoft_get_sharepoint_sites_delta` | `DRIVETOOL` | Get or exhaust a SharePoint sites delta enumeration safely. |
| `microsoft_get_sign_in_log` | `OTHERTOOL` | Get a specific sign-in log entry. |
| `microsoft_get_site` | `DRIVETOOL` | Get SharePoint site. |
| `microsoft_get_site_drive_by_id` | `DRIVETOOL` | Get a document library from a SharePoint site by drive ID. |
| `microsoft_get_site_item` | `DRIVETOOL` | Get a base item addressed through a SharePoint site. |
| `microsoft_get_site_list` | `DRIVETOOL` | Get a SharePoint site list. |
| `microsoft_get_specific_calendar_event` | `CALENDARTOOL` | Get specific calendar event. |
| `microsoft_get_subject_rights_request` | `OTHERTOOL` | Get a specific subject rights request. |
| `microsoft_get_subscription` | `OTHERTOOL` | Get a specific subscription. |
| `microsoft_get_team` | `DIRECTORYTOOL` | Get team. |
| `microsoft_get_team_channel` | `DIRECTORYTOOL` | Get team channel. |
| `microsoft_get_teams_user_activity` | `DIRECTORYTOOL` | Get Teams user activity detail report. |
| `microsoft_get_threat_intelligence_host` | `OTHERTOOL` | Get a specific threat intelligence host. |
| `microsoft_get_todo_task` | `APPSTOOL` | Get Todo task. |
| `microsoft_list_access_reviews` | `ADMINTOOL` | List access review definitions. |
| `microsoft_list_accounts` | `SYSTEMTOOL` | List accounts. |
| `microsoft_list_agreements` | `OTHERTOOL` | List agreements (terms of use). |
| `microsoft_list_applications` | `OTHERTOOL` | List app registrations. |
| `microsoft_list_booking_appointments` | `OTHERTOOL` | List booking appointments for a business. |
| `microsoft_list_booking_businesses` | `OTHERTOOL` | List booking businesses. |
| `microsoft_list_calendar_events` | `CALENDARTOOL` | List calendar events. |
| `microsoft_list_calendars` | `CALENDARTOOL` | List calendars. |
| `microsoft_list_call_records` | `OTHERTOOL` | List call records. |
| `microsoft_list_channel_message_replies` | `MAILTOOL` | List replies to a Teams channel message. |
| `microsoft_list_channel_messages` | `MAILTOOL` | List channel messages. |
| `microsoft_list_chat_message_replies` | `MAILTOOL` | List chat message replies. |
| `microsoft_list_chat_messages` | `MAILTOOL` | List chat messages. |
| `microsoft_list_chats` | `DIRECTORYTOOL` | List user chats. |
| `microsoft_list_conditional_access_policies` | `ADMINTOOL` | List conditional access policies. |
| `microsoft_list_delegated_admin_relationships` | `OTHERTOOL` | List delegated admin relationships. |
| `microsoft_list_deleted_items` | `OTHERTOOL` | List deleted directory items. |
| `microsoft_list_device_compliance_policies` | `OTHERTOOL` | List device compliance policies. |
| `microsoft_list_device_configurations` | `OTHERTOOL` | List device configurations. |
| `microsoft_list_devices` | `OTHERTOOL` | List devices registered in the directory. |
| `microsoft_list_directory_audits` | `ADMINTOOL` | List directory audit logs. |
| `microsoft_list_directory_objects` | `OTHERTOOL` | List directory objects. |
| `microsoft_list_directory_role_templates` | `ADMINTOOL` | List directory role templates. |
| `microsoft_list_directory_roles` | `ADMINTOOL` | List directory roles. |
| `microsoft_list_domain_service_configuration_records` | `ADMINTOOL` | List domain service configuration DNS records. |
| `microsoft_list_domains` | `ADMINTOOL` | List tenant domains. |
| `microsoft_list_drives` | `DRIVETOOL` | List drives. |
| `microsoft_list_education_assignments` | `OTHERTOOL` | List assignments for an education class. |
| `microsoft_list_education_classes` | `OTHERTOOL` | List education classes. |
| `microsoft_list_education_schools` | `OTHERTOOL` | List education schools. |
| `microsoft_list_education_users` | `DIRECTORYTOOL` | List education users. |
| `microsoft_list_entitlement_access_packages` | `ADMINTOOL` | List entitlement management access packages. |
| `microsoft_list_excel_tables` | `APPSTOOL` | List Excel tables. |
| `microsoft_list_excel_worksheets` | `APPSTOOL` | List Excel worksheets. |
| `microsoft_list_external_connections` | `OTHERTOOL` | List external connections. |
| `microsoft_list_file_storage_containers` | `DRIVETOOL` | List file storage containers. |
| `microsoft_list_folder_files` | `DRIVETOOL` | List folder files. |
| `microsoft_list_group_conversations` | `DIRECTORYTOOL` | List group conversations. |
| `microsoft_list_group_drives` | `DRIVETOOL` | List group drives. |
| `microsoft_list_group_members` | `DIRECTORYTOOL` | List group members. |
| `microsoft_list_group_owners` | `DIRECTORYTOOL` | List group owners. |
| `microsoft_list_groups` | `DIRECTORYTOOL` | List all Microsoft 365 groups and security groups. |
| `microsoft_list_joined_teams` | `DIRECTORYTOOL` | List joined teams. |
| `microsoft_list_learning_course_activities` | `OTHERTOOL` | List learning course activities for the current user. |
| `microsoft_list_learning_providers` | `OTHERTOOL` | List learning providers. |
| `microsoft_list_lifecycle_workflows` | `ADMINTOOL` | List lifecycle management workflows. |
| `microsoft_list_mail_attachments` | `MAILTOOL` | List attachments. |
| `microsoft_list_mail_folder_messages` | `MAILTOOL` | List messages in a specific folder. |
| `microsoft_list_mail_folders` | `MAILTOOL` | List mail folders. |
| `microsoft_list_mail_messages` | `MAILTOOL` | List mail messages. |
| `microsoft_list_managed_devices` | `OTHERTOOL` | List managed devices. |
| `microsoft_list_onenote_notebook_sections` | `APPSTOOL` | List Onenote notebook sections. |
| `microsoft_list_onenote_notebooks` | `APPSTOOL` | List notebooks owned by the current user. |
| `microsoft_list_onenote_section_pages` | `APPSTOOL` | List Onenote section pages. |
| `microsoft_list_online_meetings` | `CALENDARTOOL` | List online meetings for the current user. |
| `microsoft_list_organization` | `ADMINTOOL` | List organization properties. |
| `microsoft_list_outlook_contacts` | `CALENDARTOOL` | List Outlook contacts. |
| `microsoft_list_permission_grant_policies` | `DRIVETOOL` | List permission grant policies. |
| `microsoft_list_plan_tasks` | `APPSTOOL` | List tasks for a Planner plan. |
| `microsoft_list_planner_tasks` | `APPSTOOL` | List Planner tasks. |
| `microsoft_list_presences` | `DIRECTORYTOOL` | List presence information for users. |
| `microsoft_list_print_jobs` | `OTHERTOOL` | List print jobs for a printer. |
| `microsoft_list_print_shares` | `DRIVETOOL` | List print shares. |
| `microsoft_list_printers` | `OTHERTOOL` | List printers. |
| `microsoft_list_provisioning_logs` | `OTHERTOOL` | List provisioning logs. |
| `microsoft_list_risk_detections` | `OTHERTOOL` | List risk detections. |
| `microsoft_list_risky_users` | `DIRECTORYTOOL` | List risky users. |
| `microsoft_list_role_assignments` | `ADMINTOOL` | List role assignments. |
| `microsoft_list_role_definitions` | `ADMINTOOL` | List role definitions. |
| `microsoft_list_room_lists` | `OTHERTOOL` | List room lists. |
| `microsoft_list_rooms` | `OTHERTOOL` | List rooms. |
| `microsoft_list_secure_scores` | `OTHERTOOL` | List secure scores. |
| `microsoft_list_security_alerts` | `ADMINTOOL` | List security alerts (v2). |
| `microsoft_list_security_incidents` | `ADMINTOOL` | List security incidents. |
| `microsoft_list_sensitivity_labels` | `OTHERTOOL` | List sensitivity labels. |
| `microsoft_list_service_health` | `ADMINTOOL` | List service health overviews. |
| `microsoft_list_service_health_issues` | `ADMINTOOL` | List service health issues. |
| `microsoft_list_service_principals` | `OTHERTOOL` | List service principals. |
| `microsoft_list_service_update_messages` | `MAILTOOL` | List service update messages. |
| `microsoft_list_shared_mailbox_folder_messages` | `MAILTOOL` | List messages in a shared mailbox folder. |
| `microsoft_list_shared_mailbox_messages` | `MAILTOOL` | List messages in a shared mailbox. |
| `microsoft_list_sharepoint_site_list_items` | `DRIVETOOL` | List items in a SharePoint site list. |
| `microsoft_list_sign_in_logs` | `OTHERTOOL` | List sign-in logs. |
| `microsoft_list_site_drives` | `DRIVETOOL` | List drives for a SharePoint site. |
| `microsoft_list_site_items` | `DRIVETOOL` | List base items addressed through a SharePoint site. |
| `microsoft_list_site_lists` | `DRIVETOOL` | List lists for a SharePoint site. |
| `microsoft_list_sites` | `DRIVETOOL` | List SharePoint sites. |
| `microsoft_list_specific_calendar_events` | `CALENDARTOOL` | List events for a specific calendar. |
| `microsoft_list_subject_rights_requests` | `OTHERTOOL` | List subject rights requests. |
| `microsoft_list_subscriptions` | `OTHERTOOL` | List active webhook subscriptions. |
| `microsoft_list_team_channels` | `DIRECTORYTOOL` | List team channels. |
| `microsoft_list_team_members` | `DIRECTORYTOOL` | List team members. |
| `microsoft_list_threat_intelligence_hosts` | `OTHERTOOL` | List threat intelligence hosts. |
| `microsoft_list_todo_task_lists` | `APPSTOOL` | List Todo task lists. |
| `microsoft_list_todo_tasks` | `APPSTOOL` | List Todo tasks. |
| `microsoft_list_token_issuance_policies` | `OTHERTOOL` | List token issuance policies. |
| `microsoft_list_token_lifetime_policies` | `OTHERTOOL` | List token lifetime policies. |
| `microsoft_list_users` | `DIRECTORYTOOL` | List users. |
| `microsoft_list_virtual_events` | `CALENDARTOOL` | List virtual event townhalls. |
| `microsoft_login` | `SYSTEMTOOL` | Authenticate with Microsoft. |
| `microsoft_logout` | `SYSTEMTOOL` | Logout. |
| `microsoft_move_mail_message` | `MAILTOOL` | Move a message to a folder. |
| `microsoft_remove_application_password` | `OTHERTOOL` | Remove a password credential from an application. |
| `microsoft_remove_group_member` | `DIRECTORYTOOL` | Remove a member from a group. |
| `microsoft_reply_to_channel_message` | `MAILTOOL` | Reply to a Teams channel message. |
| `microsoft_reply_to_chat_message` | `MAILTOOL` | Reply to a chat message. |
| `microsoft_restore_deleted_item` | `OTHERTOOL` | Restore a deleted directory item. |
| `microsoft_retire_managed_device` | `OTHERTOOL` | Retire a managed device. |
| `microsoft_run_hunting_query` | `ADMINTOOL` | Run an advanced hunting query. |
| `microsoft_search_query` | `OTHERTOOL` | Search query. |
| `microsoft_search_tools` | `SYSTEMTOOL` | Search methods in this class. |
| `microsoft_send_channel_message` | `MAILTOOL` | Send channel message. |
| `microsoft_send_chat_message` | `MAILTOOL` | Send chat message. |
| `microsoft_send_mail` | `MAILTOOL` | Send mail. |
| `microsoft_send_shared_mailbox_mail` | `MAILTOOL` | Send mail from a shared mailbox. |
| `microsoft_sort_excel_range` | `APPSTOOL` | Apply a sort operation to an addressed Excel worksheet range. |
| `microsoft_start_print_job` | `OTHERTOOL` | Start an uploaded Universal Print job. |
| `microsoft_submit_print_document` | `OTHERTOOL` | Create, upload, and start one print job without exposing its upload URL. |
| `microsoft_update_admin_sharepoint` | `DRIVETOOL` | Update SharePoint admin settings. |
| `microsoft_update_application` | `OTHERTOOL` | Update an application. |
| `microsoft_update_calendar_event` | `CALENDARTOOL` | Update calendar event. |
| `microsoft_update_conditional_access_policy` | `ADMINTOOL` | Update a conditional access policy. |
| `microsoft_update_group` | `DIRECTORYTOOL` | Update a group. |
| `microsoft_update_mail_message` | `MAILTOOL` | Update a message. |
| `microsoft_update_online_meeting` | `CALENDARTOOL` | Update an online meeting. |
| `microsoft_update_org_branding` | `OTHERTOOL` | Update organization branding. |
| `microsoft_update_organization` | `ADMINTOOL` | Update organization properties. |
| `microsoft_update_outlook_contact` | `CALENDARTOOL` | Update Outlook contact. |
| `microsoft_update_place` | `OTHERTOOL` | Update a place. |
| `microsoft_update_planner_task` | `APPSTOOL` | Update Planner task. |
| `microsoft_update_planner_task_details` | `APPSTOOL` | Update Planner task details. |
| `microsoft_update_security_alert` | `ADMINTOOL` | Update a security alert (e.g. change status, assign). |
| `microsoft_update_security_incident` | `ADMINTOOL` | Update a security incident. |
| `microsoft_update_service_principal` | `OTHERTOOL` | Update a service principal. |
| `microsoft_update_specific_calendar_event` | `CALENDARTOOL` | Update specific calendar event. |
| `microsoft_update_subscription` | `OTHERTOOL` | Update/renew a subscription. |
| `microsoft_update_todo_task` | `APPSTOOL` | Update Todo task. |
| `microsoft_upload_file_content` | `DRIVETOOL` | Upload file content. |
| `microsoft_verify_domain` | `ADMINTOOL` | Verify domain ownership. |
| `microsoft_verify_login` | `SYSTEMTOOL` | Verify login status. |
| `microsoft_wipe_managed_device` | `OTHERTOOL` | Wipe a managed device. |

</details>

_79 action-routed tool(s) · 268 verbose 1:1 tool(s). Each is enabled unless its `<DOMAIN>TOOL` toggle is set false; `MCP_TOOL_MODE` selects the surface (**`intent` default** — the six verb-tools, granular set loaded on demand · `condensed` action-routed · `verbose` 1:1 · `both`). Auto-generated — do not edit._
<!-- MCP-TOOLS-TABLE:END -->

Every routed action passes the fail-closed tool policy. Reads are permitted by
default; unknown actions are treated as writes. Writes require
`MICROSOFT_ALLOW_WRITES=true`, and destructive actions additionally require
`MICROSOFT_ALLOW_DESTRUCTIVE=true`.

## MCP client configuration

<!-- MCP-CONFIG-EXAMPLES:START -->

> **Install the connector-focused `[mcp]` extra.** Examples use `microsoft-agent[mcp]` to add
> FastMCP / FastAPI through `agent-utilities[mcp]`; the required Agent Utilities core
> still carries `epistemic-graph[full]`. The `[agent-runtime]` extra additionally
> enables model orchestration.

#### stdio Transport (local IDEs — Cursor, Claude Desktop, VS Code)

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": [
        "--from",
        "microsoft-agent[mcp]",
        "microsoft-mcp"
      ],
      "env": {
        "MCP_TOOL_MODE": "intent"
      }
    }
  }
}
```

Runtime references require an alias-aware launcher such as GraphOS. Other
launchers must omit those entries and inject the resolved values through their
own runtime secret boundary.

#### Streamable-HTTP Transport (networked / production)

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": [
        "--from",
        "microsoft-agent[mcp]",
        "microsoft-mcp",
        "--transport",
        "streamable-http",
        "--port",
        "8000"
      ],
      "env": {
        "TRANSPORT": "streamable-http",
        "HOST": "127.0.0.1",
        "PORT": "8000",
        "MCP_TOOL_MODE": "intent"
      }
    }
  }
}
```

Alternatively, connect to a pre-deployed Streamable-HTTP instance by `url`:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "url": "http://localhost:8000/microsoft-mcp/mcp"
    }
  }
}
```

Run a reviewed container image as a least-privilege stdio child (no
listener or published port):

```bash
docker run -i --rm \
  --read-only \
  --cap-drop=ALL \
  --security-opt=no-new-privileges \
  --pids-limit=256 \
  --tmpfs /tmp:rw,noexec,nosuid,nodev,size=64m \
  -e TRANSPORT=stdio \
  -e MCP_TOOL_MODE=intent \
  registry.example.invalid/microsoft-agent@sha256:<digest> microsoft-mcp
```

For containerized network HTTP, supply an authenticated TLS ingress (or
direct server TLS), exact `MCP_ALLOWED_HOSTS`, and an exact trusted-proxy
CIDR policy through the operator-owned deployment profile. The generator
does not emit an unauthenticated non-loopback listener.

_Auto-generated from the code-read env surface (`MCP_TOOL_MODE` + package vars) — do not edit._
<!-- MCP-CONFIG-EXAMPLES:END -->

## A2A agent

```bash
microsoft-agent --mcp-url <mcp-url> --provider <provider> --model-id <model-id>
```

The agent uses the Agent Utilities model/configuration boundary. Provider keys,
model endpoints, Langfuse configuration, and TLS profiles are supplied externally.

## Governed ingestion

The provider contributes:

- `microsoft_agent/connectors/mcp_source_presets.json`
- `connector_manifest.yml`
- `microsoft_agent/ontology/microsoft.ttl`
- SHACL, mapping, fixture, migration, schema-fingerprint, and certification assets

Microsoft source projection persists only keyed opaque identifiers, structural node
types, and relationships. It never stores names, addresses, subjects, bodies,
filenames, URLs, timestamps, attachment bytes, or provider identifiers. The
pseudonymization key is supplied by AgentConfig or a secret store and is never
packaged or traced. Records remain quarantined until tenant, ACL, provenance,
schema, signature, and privacy requirements are satisfied. Generated signatures
and fingerprints must be regenerated whenever the tool schema or ontology changes;
stale attestations are not valid release evidence.

## Optional integrations

Power Platform, Intune, Windows companion, document roots, Office origins, and
allowlists are described by an external profile referenced with
`MICROSOFT_INTEGRATIONS_CONFIG_PATH`. The repository contains only the native
connection points and generic schemas. It does not ship organization-specific
environments, device lists, URLs, or ontologies.

## Containers

`docker/Dockerfile` has `mcp` and `agent` targets:

```bash
docker build -f docker/Dockerfile --target mcp -t <registry>/microsoft-agent:<version>-mcp .
docker build -f docker/Dockerfile --target agent -t <registry>/microsoft-agent:<version> .
```

Deployment-owned Compose or Kubernetes configuration supplies identity, trust,
storage, and network policy.

## Documentation

- [Overview](docs/overview.md)
- [Installation](docs/installation.md)
- [Authentication](docs/authentication.md)
- [Configuration](docs/configuration.md)
- [Usage](docs/usage.md)
- [Deployment](docs/deployment.md)
- [Office add-in](docs/office-addin.md)
- [Windows companion](docs/windows-companion.md)
- [Integration enrollment checklist](docs/enrollment-checklist.md)

The MkDocs navigation is defined in `mkdocs.yml`; strict documentation builds are a
release gate.

## Development

```bash
uv sync --all-extras --dev
uv run pytest
uv run ruff check .
uv run python scripts/security_sanitizer.py
uv run python scripts/security_contract.py --contract .security/security-contract.json validate
uv run python scripts/verify_api_integration.py --local
uv run mkdocs build --strict
```

See `AGENTS.md` for current-only, privacy, security, and change-discipline rules.


<!-- BEGIN agent-utilities-deployment (generated; do not edit between markers) -->

## Deploy with `agent-utilities-deployment`

Provision this package with the consolidated **`agent-utilities-deployment`**
workflow. It selects an installed-package, editable-source, or immutable-container
path; records only runtime secret and TLS-profile references in `AgentConfig`; and
runs doctor, registration, policy, observability, and rollback gates. Ask your agent
to **"deploy `microsoft-agent` with agent-utilities-deployment"**.

| Install mode | Command |
|------|---------|
| Installed package | `uv tool install "microsoft-agent[mcp]"`, then run `microsoft-mcp` |
| Editable source | `uv pip install -e ".[agent]"`, then run `microsoft-mcp` |
| Immutable container | deploy `registry.example.invalid/microsoft-agent@sha256:<digest>` through the operator-selected orchestrator |

The repository embeds no deployment profile, credential value, certificate path, or
environment-specific endpoint. Supply those at runtime through `AgentConfig` and the
configured secret provider.

<!-- END agent-utilities-deployment -->

<!-- GOVERNED-CAPABILITY:START -->
## Governed capability contract

This package ships a compact canonical skill surface with specialist procedures
kept as referenced workflows. The current MCP tools, skill metadata,
`connector_manifest.yml`, ontology, mappings, shapes, fixtures, migrations,
tool-schema fingerprints, and certification metadata form one versioned
capability contract. Validate them together; do not rely on stale tool names or
historical per-task skill wrappers.

Runtime endpoints, credentials, certificate trust, tenant identity, retention,
and observability policy are deployment inputs and are never packaged values.
See [Configuration, trust, and privacy](docs/configuration.md) before enabling a
network transport, connector ingestion, GraphOS delegation, or trace export.
<!-- GOVERNED-CAPABILITY:END -->
