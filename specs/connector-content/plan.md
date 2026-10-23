# CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT — Design and implementation plan

Governing spec: [spec.md](spec.md).

## Architecture

`pyproject.toml` declares `microsoft-mcp` as the connector identity. `create_mcp_server` resolves the packaged content of `microsoft_agent` from it.
