# Installation

`microsoft-agent` is a standard Python package and a prebuilt container image. Pick
the path that matches how you want to run it.

## Requirements

- **Python 3.11 – 3.14**.
- A registered **Microsoft Entra ID (Azure AD) application** with Microsoft Graph
  permissions, and its client id / client secret. Microsoft 365 is a managed service
  — see the [Backing Service](deployment.md#backing-service) note for configuration.

## From PyPI (recommended)

```bash
pip install microsoft-agent
```

### Optional extras

The base install is intentionally minimal. Install the extra for what you need:

| Extra | Install | Pulls in |
|---|---|---|
| `mcp` | `pip install "microsoft-agent[mcp]"` | FastMCP MCP-server runtime (`agent-utilities[mcp]`) |
| `agent` | `pip install "microsoft-agent[agent]"` | Pydantic-AI A2A agent + Logfire tracing |
| `all` | `pip install "microsoft-agent[all]"` | MCP server, agent, and Logfire — everything above |
| `test` | `pip install "microsoft-agent[test]"` | `pytest`, `pytest-asyncio`, `pytest-cov` |

```bash
# Typical: run the MCP server and the A2A agent together
pip install "microsoft-agent[all]"
```

## From source

```bash
git clone https://github.com/Knuckles-Team/microsoft-agent.git
cd microsoft-agent
pip install -e ".[all]"          # editable install with every extra
```

With [`uv`](https://docs.astral.sh/uv/):

```bash
uv pip install -e ".[all]"
uv run microsoft-mcp
```

## Prebuilt Docker image

A multi-stage, slim image is published on every release (installs
`microsoft-agent[all]`, console scripts `microsoft-mcp` and `microsoft-agent`):

```bash
docker pull knucklessg1/microsoft-agent:latest

docker run --rm -i \
  -e MICROSOFT_CLIENT_ID=your-app-registration-client-id \
  -e MICROSOFT_CLIENT_SECRET=your-client-secret \
  -e MICROSOFT_SCOPE=https://graph.microsoft.com/.default \
  knucklessg1/microsoft-agent:latest        # stdio transport (default)
```

For an HTTP server with a published port and the A2A agent, see
[Deployment](deployment.md).

## Verify the install

```bash
microsoft-mcp --help
microsoft-agent --help
python -c "import microsoft_agent; print(microsoft_agent.__version__)"
```

## Next steps

- **[Deployment](deployment.md)** — run it as a long-lived MCP server and A2A agent behind Caddy + DNS.
- **[Usage](usage.md)** — call the tools, the API, and the CLI.
- **[Configuration](deployment.md#configuration-environment)** — every environment variable.
