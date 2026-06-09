# microsoft-agent

Microsoft Graph **MCP Server + A2A Supervisor Agent** for the agent-utilities
ecosystem — manage a Microsoft 365 tenant (users, groups, calendars, mail, files,
Teams, and more) through typed, deterministic tools and natural-language delegation.

!!! info "Official documentation"
    This site is the canonical reference for `microsoft-agent`, maintained alongside
    every release.

[![PyPI](https://img.shields.io/pypi/v/microsoft-agent)](https://pypi.org/project/microsoft-agent/)
![MCP Server](https://badge.mcpx.dev?type=server 'MCP Server')
[![License](https://img.shields.io/pypi/l/microsoft-agent)](https://github.com/Knuckles-Team/microsoft-agent/blob/main/LICENSE)
[![GitHub](https://img.shields.io/badge/source-GitHub-181717?logo=github)](https://github.com/Knuckles-Team/microsoft-agent)

## Overview

`microsoft-agent` wraps the **Microsoft Graph API** with typed, deterministic MCP
tools and ships an out-of-the-box **Agent2Agent (A2A) Supervisor Agent** that
delegates work to specialized domain agents. It provides:

- **`MicrosoftGraphApi`** — a layered client over the Microsoft Graph SDK, organized
  by domain (mail, calendar, drive, directory, applications, administration), built
  on MSAL authentication.
- **Domain-routed MCP tools** — action-dispatch readers and writers across 36
  Microsoft Graph domains (users, groups, calendar, files, Teams, security, …),
  each gated by an enable flag.
- **A Supervisor-Worker A2A agent** — a confidence-gated router that classifies each
  request and engages only the relevant domain tools.

The MCP server remains inactive when credentials are absent; every domain tool set is
individually enabled or disabled by environment flag.

## Explore the documentation

<div class="grid cards" markdown>

- :material-rocket-launch: **[Installation](installation.md)** — pip, source, extras, and the prebuilt Docker image.
- :material-server-network: **[Deployment](deployment.md)** — run the MCP and agent servers, Docker Compose, Caddy + Technitium.
- :material-console: **[Usage](usage.md)** — the MCP tools, the `MicrosoftGraphApi` client, and the CLI.
- :material-sitemap: **[Overview](overview.md)** — capabilities, tool surface, and the agent architecture.
- :material-tag-multiple: **[Concepts](concepts.md)** — the `CONCEPT:MSFT-*` domain registry.

</div>

## Quick start

```bash
pip install "microsoft-agent[mcp]"
microsoft-mcp                    # stdio MCP server (default transport)
```

Connect it to a Microsoft 365 tenant:

```bash
export MICROSOFT_CLIENT_ID=your-app-registration-client-id
export MICROSOFT_CLIENT_SECRET=your-client-secret
export MICROSOFT_SCOPE=https://graph.microsoft.com/.default
microsoft-mcp --transport http --host 0.0.0.0 --port 8000
```

See **[Installation](installation.md)** and **[Deployment](deployment.md)** for the
full matrix (PyPI extras, Docker image, all transports, the agent server, reverse
proxy, DNS).
