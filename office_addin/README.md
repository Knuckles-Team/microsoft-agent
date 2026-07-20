# Microsoft Agent Office Companion

This directory contains one Office.js task-pane add-in that runs in both Word and PowerPoint. It performs live, user-visible edits in the open Office document and provides a guarded connection point for the `microsoft-agent` backend.

The add-in does not contain an Entra client secret, Microsoft Graph credential, or backend access token. A token entered in the task pane is held only in page memory and is lost when the pane reloads. Authentication can be connected later without changing the Office document operations.

## Features and requirement sets

| Feature | Host | Minimum runtime requirement |
| --- | --- | --- |
| Report host, platform, Office version, and capabilities | Word or PowerPoint | Office.js common API |
| Read the current selection | Word | WordApi 1.1 |
| Replace, prepend to, or append to the current selection | Word | WordApi 1.1 |
| Replace every literal placeholder in the document body | Word | WordApi 1.1 |
| List, append, and delete slides | PowerPoint | PowerPointApi 1.2 |
| Add a positioned text box | PowerPoint | PowerPointApi 1.4 |
| Call the configured backend health endpoint | Either | HTTPS and backend CORS support |
| Pair the open window for typed agent commands | Either | HTTPS, exact-origin CORS, and a one-time pairing secret |

The XML manifest intentionally does not impose a host-specific requirement set globally. Requiring `WordApi` in a manifest shared with PowerPoint (or vice versa) would make the add-in unavailable in the other host. Each operation checks its requirement set at runtime and reports a clear unsupported-requirement error.

PowerPoint slide numbers in the UI are one-based. Office.js slide collection indexes are zero-based; the adapter performs and validates that conversion. A presentation's final remaining slide cannot be deleted through this UI.

## Development setup

Prerequisites:

- Node.js 22.15 or newer.
- Python 3.11 or newer for the bounded, offline manifest contract check.
- A Microsoft 365 build of Word and/or PowerPoint that supports Office web add-ins.
- On first development-server launch, permission to install the local HTTPS development certificate.

From this directory:

```bash
npm install
npm run check
npm run dev
```

The development server listens at `https://localhost:3000`. `office-addin-dev-certs` creates/trusts the development certificate used by webpack. The build output is written to `dist/`.

Useful individual commands:

```bash
npm run typecheck
npm test
npm run build
npm run validate:manifest
```

`npm run validate:manifest` performs a deterministic offline XML and provider-contract check. Microsoft 365 performs the authoritative platform validation when the deployment-owned manifest is uploaded through Integrated Apps or sideloading.

## Sideloading

1. Start the HTTPS server with `npm run dev`.
2. Open a document in Word or a presentation in PowerPoint.
3. Open **Home > Add-ins > More Add-ins** (the exact label varies by Office build).
4. Choose **Upload My Add-in** and upload `manifest.xml`.
5. Open **Microsoft Agent Office Companion** from the add-ins UI if the pane does not open automatically.

For Office on the web, upload the manifest from the add-ins dialog in the open document. For desktop Windows testing at organization scale, a trusted network-share catalog can host the manifest. Production deployments should use the Microsoft 365 admin center's Integrated Apps workflow rather than localhost.

Microsoft's current sideloading instructions are at:

- <https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing>
- <https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins>

## Configure the backend

Edit `public/config.json` before building or deploying:

```json
{
  "allowedBackendOrigins": [
    "https://microsoft-agent.example.com"
  ],
  "defaultBackendUrl": "https://microsoft-agent.example.com",
  "healthPath": "/health",
  "requestTimeoutMs": 10000
}
```

Only complete HTTPS origins are accepted. Paths, embedded credentials, query strings, fragments, HTTP, wildcard hosts, suffix matches, and URLs absent from `allowedBackendOrigins` are rejected. The backend URL can be changed in the task pane only to another exact origin in that deployment list. The approved URL may be saved in local storage; tokens never are.

Also replace the development-only `connect-src 'self' https://localhost:*` directive in `public/taskpane.html` with the same explicit production origins. Set an equivalent `Content-Security-Policy` HTTP response header in production, including an Office-compatible `frame-ancestors` policy; response headers provide stronger enforcement than an HTML meta policy, and `frame-ancestors` is not honored in a meta element. Do not add a broad `connect-src https:` rule.

The health endpoint contract is:

- `GET` the configured `healthPath`.
- Return a successful status with `Content-Type: application/json` (or a `+json` type).
- Keep the response at or below 1 MiB.
- If authentication is enabled, accept `Authorization: Bearer <token>`.
- Permit CORS only from the exact origins that host this task pane, allow the `Authorization` request header, and do not require browser cookies.

Requests omit cookies, reject redirects, use no referrer, enforce a 1–60 second timeout, constrain request/response sizes, and do not display error-response bodies. These controls prevent an entered token from being redirected or sent to an origin not approved by deployment configuration.

The Python MCP service exposes `/health` and the `/office-bridge/*` routes. The default `https://localhost:8000` entry still requires an HTTPS listener/certificate and `MICROSOFT_OFFICE_ADDIN_ORIGINS=https://localhost:3000`; a plain HTTP development server is deliberately rejected.

## Pair with the agent

The MCP `create_office_pairing` tool returns a random, one-time secret bound to
either Word or PowerPoint and a human-readable label. Paste that secret into
**Agent pairing** within five minutes. The task pane exchanges it for an
eight-hour session credential, clears the input, holds the credential only in
page memory, and begins long-polling typed commands. Use **Disconnect** to
revoke the session immediately; reloading loses the browser credential and the
server revokes the abandoned session after its idle timeout.

Supported remote commands exactly match the visible operations listed above.
There is no command for arbitrary JavaScript, Office Scripts, macros, file
paths, or backend URLs. The server delivers one command at a time without
automatic mutation redelivery, caps pairings/sessions/queues, expires commands,
validates command-specific results, and binds every command/result to one
session and host. Slide deletion remains subject to the MCP destructive-action
policy and also requires confirmation in the paired task pane.

## Configure production identity

The pairing bridge uses a short-lived, audience-restricted backend session. The deployment must establish that session through one of these identity flows before enabling administrative access to the MCP service:

1. Add Office single sign-on with a registered Entra application and exchange the Office identity token at the backend. The manifest then needs the deployment's non-secret application ID and resource URI in `WebApplicationInfo`.
2. Use MSAL in the task pane with authorization code + PKCE. Public-client IDs are identifiers, not secrets; never put a client secret in browser code.
3. Let the backend issue a short-lived, audience-restricted session token after an authenticated bootstrap.

Whichever approach is selected, inject the token through `BackendClient`; do not persist it in local storage and do not broaden the origin allowlist. Validate issuer, audience, tenant, scopes/roles, nonce, and expiry on the backend.

## Production deployment checklist

- Change every `https://localhost:3000` URL in `manifest.xml` to the final HTTPS task-pane origin.
- Replace `public/config.json` backend origins and the HTML `connect-src` policy with exact production origins.
- Serve `dist/` over HTTPS with CSP, `X-Content-Type-Options: nosniff`, a restrictive `Permissions-Policy`, and appropriate caching rules. Do not cache user-specific tokens or responses.
- Configure backend CORS for the task-pane origin and test browser preflight requests carrying `Authorization`.
- Replace the minimal generated development PNG with marketplace-compliant 32×32 and 64×64 PNG artwork if distributing through Microsoft Marketplace. The SVG remains the task-pane UI icon; `icon.png.b64` is decoded by webpack solely to provide a valid internal-sideload manifest asset.
- Run `npm run check`, then exercise every mutation on a disposable document and presentation in each supported Office platform.
- Confirm the tenant's policy permits Office add-ins and deploy the manifest using Integrated Apps.
- Add user confirmation and backend audit events for any future agent-triggered operation that changes document content.

## Source layout

```text
office_addin/
├── manifest.xml
├── public/
│   ├── assets/icon.svg
│   ├── assets/icon.png.b64
│   ├── config.json
│   ├── styles.css
│   └── taskpane.html
├── src/
│   ├── backend-client.ts
│   ├── office-bridge.ts
│   ├── office-operations.ts
│   ├── taskpane.ts
│   └── validation.ts
└── tests/
    ├── backend-client.test.ts
    ├── office-bridge.test.ts
    └── validation.test.ts
```

The Office host adapter is isolated from DOM code, and the backend client/validation code has no dependency on Office.js. This keeps security-sensitive URL logic unit-testable without launching Office.

## API references

- Runtime requirement checks: <https://learn.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements>
- Word JavaScript API: <https://learn.microsoft.com/javascript/api/word>
- Add and delete slides: <https://learn.microsoft.com/office/dev/add-ins/powerpoint/add-slides>
- PowerPoint text boxes (`PowerPointApi 1.4`): <https://learn.microsoft.com/javascript/api/powerpoint/powerpoint.shapecollection>
