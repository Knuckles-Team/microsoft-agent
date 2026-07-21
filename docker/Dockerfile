# syntax=docker/dockerfile:1.7

# The digest pins the multi-platform Python 3.12 slim image. Refresh it only as
# part of an audited dependency update.
ARG PYTHON_IMAGE=python:3.12-slim@sha256:57cd7c3a7a273101a6485ba99423ee568157882804b1124b4dd04266317710de

FROM ${PYTHON_IMAGE} AS source
WORKDIR /build
COPY pyproject.toml README.md LICENSE MANIFEST.in ./
COPY microsoft_agent ./microsoft_agent

FROM source AS builder-mcp
RUN python -m venv /opt/microsoft-agent \
    && /opt/microsoft-agent/bin/python -m pip install \
        --disable-pip-version-check --no-cache-dir ".[mcp]"

FROM source AS builder-agent
RUN python -m venv /opt/microsoft-agent \
    && /opt/microsoft-agent/bin/python -m pip install \
        --disable-pip-version-check --no-cache-dir ".[agent]"

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

FROM runtime-base AS agent
COPY --from=builder-agent --chown=65532:65532 /opt/microsoft-agent /opt/microsoft-agent
ENTRYPOINT ["microsoft-agent"]
