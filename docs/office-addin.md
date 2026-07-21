# Word and PowerPoint add-in

`office_addin/` is a shared Office.js task pane for desktop/web Word and
PowerPoint. It is intentionally complementary to Graph: Graph stores and moves
the files, while Office.js edits the document currently open in the user's
Office host.

Implemented Word operations:

- report current host and requirement-set support;
- read the current selection;
- insert or replace selection text;
- replace literal placeholders throughout the document.

Implemented PowerPoint operations:

- report current host and requirement-set support;
- list slides;
- add and delete slides;
- add text boxes when `PowerPointApi 1.4` is available.

## Pair an open document with the agent

The Python MCP server and task pane include a typed bridge for those same live
operations. It does not accept JavaScript, Office scripts, caller-selected
URLs, or an open-ended action name.

1. Start the MCP server over HTTPS and set
   `MICROSOFT_OFFICE_ADDIN_ORIGINS` to the task pane's exact HTTPS origin.
2. Enable `documents` and `MICROSOFT_ALLOW_WRITES`. Enable
   `MICROSOFT_ALLOW_DESTRUCTIVE` only if the agent may delete slides.
3. Ask the agent to call `create_office_pairing` with `Word` or `PowerPoint`
   and a recognizable window label.
4. Paste the returned one-time secret into **Agent pairing** in that intended
   task pane. The secret expires after five minutes and cannot be reused.
5. Use `list_office_sessions` to obtain the non-secret session ID, then call a
   host-specific live Office tool. A tool waits up to its requested timeout;
   use `get_office_command_result` if it returns a queued or delivered state.

The task pane long-polls one command at a time and posts a discriminated,
validated result. Pairing/session credentials are held only as SHA-256 digests
by Python and only in page memory by the add-in. Sessions expire after eight
hours, after 15 minutes without polling, or immediately on a successful
disconnect. Commands expire after two minutes and result retention is bounded.
The default store is process-local and fail-closed on restart, so run the HTTP
MCP transport with one worker. A shared store is required before horizontal
scaling.

Remote slide deletion also displays a native confirmation in the paired task
pane; declining it returns a typed failure and leaves the presentation intact.

The task pane uses a strict Content Security Policy and an exact HTTPS backend
origin allowlist from `config.json`. A bearer token, when supplied, remains in
memory. Requests have time and response-size limits, do not follow redirects,
and never accept a caller-selected origin. Bridge routes additionally reject
requests without an exact configured `Origin`, keep CORS credentials disabled,
and require the short-lived session bearer for polling and results.

Development:

```bash
cd office_addin
npm ci
npm test
npm run typecheck
npm run build
```

Use the add-in's README for certificate and sideload instructions. Production
deployment should use Microsoft 365 integrated-app deployment and an HTTPS
origin registered in both the manifest and `config.json`. The backend health
contract is `GET /health`, JSON no larger than 1 MiB, optional bearer token, and
exact-origin CORS. Local HTTP is deliberately not accepted by the add-in.

Office.js feature support differs by host/build. The UI reports missing
requirement sets and disables only the unsupported operation rather than
pretending the edit succeeded. Headless document creation remains available via
the Python OOXML service when no Office host is open.
