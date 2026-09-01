"""Static supply-chain contract for provider container targets."""

import re
import tomllib
from pathlib import Path

from packaging.specifiers import SpecifierSet
from packaging.version import Version

DOCKERFILE = Path(__file__).resolve().parents[1] / "docker" / "Dockerfile"
PYPROJECT = DOCKERFILE.parents[1] / "pyproject.toml"


def test_container_builds_local_source_from_a_digest_pinned_base() -> None:
    content = DOCKERFILE.read_text(encoding="utf-8")

    image = re.search(
        r"^ARG PYTHON_IMAGE=python:(\d+\.\d+)-slim@sha256:([0-9a-f]{64})$",
        content,
        re.MULTILINE,
    )
    assert image is not None
    project = tomllib.loads(PYPROJECT.read_text(encoding="utf-8"))["project"]
    assert Version(image.group(1)) in SpecifierSet(project["requires-python"])
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
