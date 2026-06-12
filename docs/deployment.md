# Deployment

<!-- BEGIN GENERATED: deployment-options -->
## Deployment Options

`microsoft-agent` exposes its MCP server (console script `microsoft-mcp`) four ways. Pick the row that
matches where the server runs relative to your MCP client, then copy the matching
`mcp_config.json` below. Replace the `<your-…>` placeholders with the values from the **Configuration / Environment Variables** section.

| # | Option | Transport | Where it runs | `mcp_config.json` key |
|---|--------|-----------|---------------|------------------------|
| 1 | stdio | `stdio` | client launches a subprocess | `command` |
| 2 | Streamable-HTTP (local) | `streamable-http` | a local network port | `command` or `url` |
| 3 | Local container / uv | `stdio` or `streamable-http` | Docker / Podman / uv on this host | `command` or `url` |
| 4 | Remote URL | `streamable-http` | a remote host behind Caddy | `url` |

### 1. stdio (local subprocess)

The client launches the server over stdio via `uvx` — best for local IDEs
(Cursor, Claude Desktop, VS Code):

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": ["--from", "microsoft-agent", "microsoft-mcp"],
      "env": {
        "OIDC_CONFIG_URL": "<your-oidc_config_url>",
        "OIDC_BASE_URL": "<your-oidc_base_url>",
        "LLM_BASE_URL": "<your-llm_base_url>"
      }
    }
  }
}
```

### 2. Streamable-HTTP (local process)

Run the server as a long-lived HTTP process:

```bash
uvx --from microsoft-agent microsoft-mcp --transport streamable-http --host 0.0.0.0 --port 8000
curl -s http://localhost:8000/health        # {"status":"OK"}
```

Then either let the client launch it:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "uvx",
      "args": ["--from", "microsoft-agent", "microsoft-mcp", "--transport", "streamable-http", "--port", "8000"],
      "env": {
        "TRANSPORT": "streamable-http",
        "HOST": "0.0.0.0",
        "PORT": "8000",
        "OIDC_CONFIG_URL": "<your-oidc_config_url>",
        "OIDC_BASE_URL": "<your-oidc_base_url>",
        "LLM_BASE_URL": "<your-llm_base_url>"
      }
    }
  }
}
```

…or connect to the already-running process by URL:

```json
{
  "mcpServers": {
    "microsoft-mcp": { "url": "http://localhost:8000/mcp" }
  }
}
```

### 3. Local container / uv

**(a) Launch a container directly from `mcp_config.json`** (stdio over the container —
no ports to manage). Swap `docker` for `podman` for a daemonless runtime:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-e", "TRANSPORT=stdio",
        "-e", "OIDC_CONFIG_URL=<your-oidc_config_url>",
        "-e", "OIDC_BASE_URL=<your-oidc_base_url>",
        "-e", "LLM_BASE_URL=<your-llm_base_url>",
        "knucklessg1/microsoft-agent:latest"
      ]
    }
  }
}
```

**(b) Run a local streamable-http container, then connect by URL:**

```bash
docker run -d --name microsoft-mcp -p 8000:8000 \
  -e TRANSPORT=streamable-http \
  -e PORT=8000 \
  -e OIDC_CONFIG_URL="<your-oidc_config_url>" \
  -e OIDC_BASE_URL="<your-oidc_base_url>" \
  -e LLM_BASE_URL="<your-llm_base_url>" \
  knucklessg1/microsoft-agent:latest
# or, from a clone of this repo:
docker compose -f docker/mcp.compose.yml up -d
```

```json
{
  "mcpServers": {
    "microsoft-mcp": { "url": "http://localhost:8000/mcp" }
  }
}
```

**(c) From a local checkout with `uv`:**

```bash
uv run microsoft-mcp --transport streamable-http --port 8000
```

### 4. Remote URL (deployed behind Caddy)

When the server is deployed remotely (e.g. as a Docker service) and published through
Caddy on the internal `*.arpa` zone, connect with the `"url"` key — no local process or
image required:

```json
{
  "mcpServers": {
    "microsoft-mcp": { "url": "http://microsoft-mcp.arpa/mcp" }
  }
}
```

Caddy reverse-proxies `http://microsoft-mcp.arpa` to the container's `:8000`
streamable-http listener; `http://microsoft-mcp.arpa/health` returns
`{"status":"OK"}` when the service is live.
<!-- END GENERATED: deployment-options -->

This page covers running `microsoft-agent` as a long-lived service: the MCP
transports, the A2A agent server, a Docker Compose stack, putting it behind a Caddy
reverse proxy, and giving it a DNS name with Technitium.

> `microsoft-agent` ships **two console scripts**: an **MCP server** (`microsoft-mcp`)
> exposing the Microsoft Graph tool surface, and an **A2A agent server**
> (`microsoft-agent`) — a Supervisor-Worker agent that consumes those tools over the
> network. The agent connects to the MCP server via `MCP_URL`.

## Run the MCP server

The transport is selected with `--transport` (or the `TRANSPORT` env var):

=== "stdio (default)"

    ```bash
    microsoft-mcp
    ```
    For IDE / desktop MCP clients that launch the server as a subprocess.

=== "streamable-http"

    ```bash
    microsoft-mcp --transport http --host 0.0.0.0 --port 8000
    ```
    A network server with a `/health` endpoint and `/mcp` route.

=== "sse"

    ```bash
    microsoft-mcp --transport sse --host 0.0.0.0 --port 8000
    ```

Health check (HTTP transports):

```bash
curl -s http://localhost:8000/health        # {"status":"OK"}
```

## Configuration (environment)

`microsoft-agent` is configured entirely from the environment. The **required** set
for the Microsoft Graph connection:

| Var | Default | Meaning |
|---|---|---|
| `MICROSOFT_HOST` | `https://graph.microsoft.com` | Microsoft Graph base URL |
| `MICROSOFT_CLIENT_ID` | — | Entra ID application (client) id |
| `MICROSOFT_CLIENT_SECRET` | — | Client secret for the app registration |
| `MICROSOFT_SCOPE` | `https://graph.microsoft.com/.default` | OAuth scope |
| `MICROSOFT_GRANT_TYPE` | `client_credentials` | OAuth grant type |
| `MICROSOFT_TOKEN` | — | Direct user bearer token (instead of client credentials) |

Plus `HOST` / `PORT` / `TRANSPORT` for HTTP transports. Each of the 36 Microsoft
Graph domains is gated by an individual `*TOOL` flag (for example `MAILTOOL`,
`CALENDARTOOL`, `USERTOOL`) so the registered tool surface can be scoped per
deployment. The full set, including the OIDC / A2A authentication and observability
variables, is documented in
[`.env.example`](https://github.com/Knuckles-Team/microsoft-agent/blob/main/.env.example).
Copy it to `.env` and populate only the values you use.

### Backing Service

Microsoft 365 and the Microsoft Graph API are a **managed, software-as-a-service
platform** operated by Microsoft — there is no self-hosted backing system to deploy.
Connecting `microsoft-agent` requires only configuration: register an application in
**Microsoft Entra ID (Azure AD)**, grant it the Microsoft Graph permissions your
workloads need, and provide the client id, client secret, and scope through the
environment variables above. The server remains inactive when those credentials are
absent.

## Docker Compose

The repo ships [`docker/mcp.compose.yml`](https://github.com/Knuckles-Team/microsoft-agent/blob/main/docker/mcp.compose.yml),
which runs both the MCP server and the A2A agent. It reads a sibling `.env` and
publishes the HTTP MCP server alongside the agent:

```yaml
services:
  microsoft-mcp:
    image: knucklessg1/microsoft-agent:latest
    container_name: microsoft-mcp
    hostname: microsoft-mcp
    command: ["microsoft-mcp"]
    restart: always
    env_file:
      - .env
    environment:
      - PYTHONUNBUFFERED=1
      - HOST=0.0.0.0
      - PORT=8000
      - TRANSPORT=streamable-http
    ports:
      - "8000:8000"
    healthcheck:
      test: ["CMD", "python3", "-c", "import urllib.request; urllib.request.urlopen('http://localhost:8000/health')"]
      interval: 30s
      timeout: 10s
      retries: 3

  microsoft-agent:
    image: knucklessg1/microsoft-agent:latest
    container_name: microsoft-agent
    hostname: microsoft-agent
    command: ["microsoft-agent"]
    depends_on:
      - microsoft-mcp
    restart: always
    env_file:
      - .env
    environment:
      - HOST=0.0.0.0
      - PORT=9000
      - MCP_URL=http://microsoft-mcp:8000/mcp
      - PROVIDER=openai
      - MODEL_ID=gpt-4o
    ports:
      - "9000:9000"
```

```bash
cp .env.example .env          # then edit MICROSOFT_* values
docker compose -f docker/mcp.compose.yml up -d
docker compose -f docker/mcp.compose.yml logs -f
```

## Run the A2A agent server

The A2A Supervisor Agent (`microsoft-agent`) is a separate server that consumes the
MCP tool surface over the network. Point it at the running MCP server with `--mcp-url`:

```bash
microsoft-agent \
  --provider openai \
  --model-id gpt-4o \
  --api-key sk-... \
  --mcp-url http://localhost:8000/mcp \
  --host 0.0.0.0 \
  --port 9000
```

It exposes the following endpoints (default port `9000`):

| Endpoint | Path | Purpose |
|---|---|---|
| Web UI | `http://localhost:9000/` | Interactive console (when `--web` is enabled) |
| A2A | `http://localhost:9000/a2a` | Agent2Agent protocol (discovery: `/a2a/.well-known/agent.json`) |
| AG-UI | `http://localhost:9000/ag-ui` | AG-UI streaming endpoint (POST) |

The container build also runs the agent server from
[`docker/compose.yml`](https://github.com/Knuckles-Team/microsoft-agent/blob/main/docker/compose.yml).
The agent reads `MCP_URL`, `PROVIDER`, `MODEL_ID`, `LLM_BASE_URL`, and `LLM_API_KEY`
from the environment.

## Behind a Caddy reverse proxy

Expose the HTTP server on a hostname with automatic TLS. Add to your `Caddyfile`:

```caddy
# Internal (self-signed) — homelab .arpa zone
microsoft-agent.arpa {
    tls internal
    reverse_proxy microsoft-mcp:8000
}
```

```caddy
# Public — automatic Let's Encrypt
microsoft-agent.example.com {
    reverse_proxy microsoft-mcp:8000
}
```

Reload Caddy:

```bash
docker compose -f services/caddy/compose.yml exec caddy caddy reload --config /etc/caddy/Caddyfile
```

## DNS with Technitium

Point the hostname at the host running Caddy. Via the Technitium API:

```bash
curl -s "http://technitium.arpa:5380/api/zones/records/add" \
  --data-urlencode "token=$TECHNITIUM_DNS_TOKEN" \
  --data-urlencode "domain=microsoft-agent.arpa" \
  --data-urlencode "zone=arpa" \
  --data-urlencode "type=A" \
  --data-urlencode "ipAddress=10.0.0.10" \
  --data-urlencode "ttl=3600"
```

…or add an **A record** `microsoft-agent.arpa → <caddy-host-ip>` in the Technitium web
console (`http://technitium.arpa:5380`). The ecosystem
[`technitium-dns-mcp`](https://knuckles-team.github.io/technitium-dns-mcp/) automates
this as a tool.

## Register with an MCP client

Add to your client's `mcp_config.json` (multiplexer nickname `ms`):

```json
{
  "mcpServers": {
    "microsoft-agent": {
      "command": "uv",
      "args": ["run", "microsoft-mcp"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your-app-registration-client-id",
        "MICROSOFT_CLIENT_SECRET": "your-client-secret",
        "MICROSOFT_SCOPE": "https://graph.microsoft.com/.default",
        "MICROSOFT_GRANT_TYPE": "client_credentials"
      }
    }
  }
}
```

For a remote HTTP server, point the client at `http://microsoft-agent.arpa/mcp` instead.
