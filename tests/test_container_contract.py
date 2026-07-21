"""Static supply-chain contract for provider container targets."""

from pathlib import Path

DOCKERFILE = Path(__file__).resolve().parents[1] / "docker" / "Dockerfile"


def test_container_builds_local_source_from_a_digest_pinned_base() -> None:
    content = DOCKERFILE.read_text(encoding="utf-8")

    assert "python:3.12-slim@sha256:" in content
    assert "COPY pyproject.toml README.md LICENSE MANIFEST.in ./" in content
    assert '".[mcp]"' in content
    assert '".[agent]"' in content
    assert "microsoft-agent[mcp]>=" not in content
    assert "microsoft-agent[agent]>=" not in content
    assert "--prerelease" not in content
    assert "ghcr.io/" not in content


def test_container_runtime_is_unprivileged_and_has_no_auth_bypass_default() -> None:
    content = DOCKERFILE.read_text(encoding="utf-8")

    assert "USER 65532:65532" in content
    assert "--no-create-home" in content
    assert "AUTH_TYPE" not in content
    assert 'ENTRYPOINT ["microsoft-mcp"]' in content
    assert 'ENTRYPOINT ["microsoft-agent"]' in content
