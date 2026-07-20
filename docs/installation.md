# Installation

Microsoft Agent supports Python 3.11 through 3.14. Install only the capability
groups required by the deployment.

## Package extras

| Extra | Purpose |
|---|---|
| `mcp` | Microsoft Graph MCP server and Agent Utilities MCP runtime |
| `agent` | optional Agent Utilities A2A agent runtime |
| `documents` | Word and PowerPoint document generation |
| `cloud` | Azure managed and workload identity support |
| `windows` | outbound Windows companion client and service |
| `control-plane` | authenticated Windows control-plane service |
| `all` | every supported runtime capability |
| `test` | repository test dependencies |

Install the MCP server:

```bash
python -m pip install "microsoft-agent[mcp]"
```

Install a broader approved profile when required:

```bash
python -m pip install "microsoft-agent[agent,documents,cloud]"
```

Agent Utilities resolves Epistemic Graph with the `full` feature set, including the
folded numeric ABI. Do not substitute a minimal or numeric-only engine artifact.

## Source checkout

For development:

```bash
git clone <repository-url> microsoft-agent
cd microsoft-agent
uv sync --all-extras --dev
```

Keep the package cache, virtual environment, build output, and runtime data on a
filesystem with adequate capacity. Source checkouts must not contain runtime token
caches, databases, endpoint profiles, credentials, or generated operator state.

## Containers

Build the dedicated targets from the checked-in Dockerfile:

```bash
docker build -f docker/Dockerfile --target mcp -t <registry>/microsoft-agent:<version>-mcp .
docker build -f docker/Dockerfile --target agent -t <registry>/microsoft-agent:<version> .
```

Images contain application code only. Supply identity, TLS trust, secret
references, configuration, and writable storage at deployment time.

## Verify the installation

```bash
microsoft-mcp --help
microsoft-agent --help
python -c "import microsoft_agent; print(microsoft_agent.__version__)"
```

Then run the Agent Utilities doctor against the deployment-owned AgentConfig
profile. Installation is not release-ready until the doctor confirms engine,
identity, trust, permission, observability, and optional dependency readiness
without exposing values.

Continue with [Configuration](configuration.md),
[Authentication](authentication.md), and [Deployment](deployment.md).
