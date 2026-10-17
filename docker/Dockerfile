# syntax=docker/dockerfile:1.7
# MCP-only image: ships the MCP server and nothing else (no agent mode).
#   docker build --target mcp -t <name>:mcp .
FROM ${PYTHON_IMAGE} AS source
WORKDIR /build
COPY pyproject.toml README.md LICENSE MANIFEST.in ./
COPY microsoft_agent ./microsoft_agent

FROM source AS builder-mcp
RUN python -m venv /opt/microsoft-agent \
    && /opt/microsoft-agent/bin/python -m pip install \
        --disable-pip-version-check --no-cache-dir ".[mcp]"

FROM ${PYTHON_IMAGE} AS runtime-base
RUN groupadd --system --gid 65532 microsoft-agent \
    && useradd --system --uid 65532 --gid 65532 --no-create-home \
        --shell /usr/sbin/nologin microsoft-agent
ENV PATH="/opt/microsoft-agent/bin:${PATH}" \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1
USER 65532:65532

FROM runtime-base AS mcp
COPY --from=builder-mcp --chown=65532:65532 /opt/microsoft-agent /opt/microsoft-agent
ENTRYPOINT ["microsoft-mcp"]
