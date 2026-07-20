# Authentication handoff checklist

Complete this checklist when tenant authentication is available. No secrets are
needed for build or test.

1. Create or select the single-tenant Entra application. The helper under
   `deployment/entra` can create one without a secret.
2. Add delegated profiles `productivity,collaboration`; add `device_read` or
   `device_admin` only if Intune is required. Review every resolved permission.
3. Grant administrator consent where the tenant requires it. Separately grant
   `Sites.Selected` access to each approved SharePoint site.
4. Choose broker/browser for a person, certificate/application for a daemon, or
   managed/workload identity on Azure. Do not place a secret in Git or a command
   line.
5. Populate a local `.env` from `.env.example`. Start with writes and
   destructive actions disabled.
6. Copy the integration example to a protected local file. Replace Dataverse,
   named-flow, Intune-device, and companion-device placeholders; keep unused
   integrations absent or disabled.
7. Run `microsoft-config --require-identity`, then
   `microsoft-mcp --transport stdio`; call `get_microsoft_configuration`,
   then `login` for delegated mode and `verify_login`. If configured, call
   `login_power_platform` and `login_windows_companion` to grant the separate
   resource-audience consents.
8. Verify one read in each consented service: latest mail, calendar, a drive,
   selected SharePoint site, Teams, Dataverse flow list, and Intune device.
9. Enable `MICROSOFT_ALLOW_WRITES=true`; test a draft email, test calendar
   event, generated document upload, and a non-production named flow.
10. Deploy the Office add-in over HTTPS and verify host capability reporting
    before testing edits.
11. Enroll one non-production laptop, confirm outbound relay health, and test
    inventory before any sensitive or state-changing action.
12. Enable destructive actions only after approval/audit capture is working.
    Test against a designated lab device, then disable the flag again unless it
    is operationally required.

Record tenant/client IDs, consent date, profile names, site grants, allowed
device IDs, certificate expiry, and test correlation IDs. Do not record tokens,
private keys, trigger URLs, mail/document content, or clipboard/file payloads.
