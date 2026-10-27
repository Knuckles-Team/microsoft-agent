"""Keep the connector's canonical Pages and packaged skill contract explicit."""

from __future__ import annotations

import json
import re
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[1]


def _workflow(name: str):
    return yaml.safe_load((ROOT / ".github/workflows" / name).read_text())


def _actions_on(workflow: dict):
    # PyYAML's YAML 1.1 reader represents the Actions ``on`` key as True.
    return workflow["on"] if "on" in workflow else workflow[True]


def test_pages_publishes_spec_dashboard_and_supports_manual_refresh():
    workflow = _workflow("pages.yml")
    publisher = workflow["jobs"]["publish-pages"]
    actions = _actions_on(workflow)

    assert publisher["uses"] == (
        "Knuckles-Team/pipelines/.github/workflows/pages_pipeline.yml@main"
    )
    assert publisher["with"] == {
        "content_source": "docs",
        "spec_dashboard_enabled": True,
        "spec_dashboard_config": ".config/spec-dashboard.json",
    }
    assert "workflow_dispatch" in actions
    assert {
        "docs/**",
        "mkdocs.yml",
        "README.md",
        "specs/**",
        ".config/spec-dashboard.json",
        ".github/workflows/pages.yml",
    } <= set(actions["push"]["paths"])

    config = json.loads((ROOT / ".config/spec-dashboard.json").read_text())
    assert config == {
        "version": 1,
        "sources": [{"glob": "specs/*/status.json", "kind": "spec"}],
    }


def test_single_pages_publisher_has_no_legacy_docs_workflow():
    assert not (ROOT / ".github/workflows/docs.yml").exists()


def test_development_skill_names_actual_runtime_and_focused_tests():
    skill = ROOT / "microsoft_agent/skills/microsoft-agent-development/SKILL.md"
    content = skill.read_text()

    assert "name: microsoft-agent-development" in content
    assert "## Rapid delivery: the 5-minute contract" in content
    for path in (
        "microsoft_agent/__init__.py",
        "microsoft_agent/auth.py",
        "microsoft_agent/api_client.py",
        "microsoft_agent/mcp_server.py",
        "tests/test_publication_workflow.py",
        "tests/test_default_content_surfaces.py",
        "tests/test_startup.py",
        "tests/test_mcp_server.py",
        "tests/test_auth.py",
    ):
        assert (ROOT / path).exists(), path
    for runtime_name in (
        "agent-connector-sdk",
        "MICROSOFT_",
        "Microsoft Graph",
        "ChangeEnvelope",
        "TLS",
    ):
        assert runtime_name in content


def test_packaged_skill_and_version_contract_are_coherent():
    version_text = (ROOT / "microsoft_agent/_version.py").read_text()
    version = re.search(r'__version__ = "([^"]+)"', version_text)
    assert version is not None
    lock = (ROOT / "uv.lock").read_text()
    assert 'name = "microsoft-agent"' in lock
    package_data = (ROOT / "pyproject.toml").read_text()
    assert 'microsoft_agent = ["connectors/**", "ontology/**", "prompts/**", "skills/**"]' in package_data
    assert (
        ROOT / "microsoft_agent/skills/microsoft-agent-development/SKILL.md"
    ).is_file()
