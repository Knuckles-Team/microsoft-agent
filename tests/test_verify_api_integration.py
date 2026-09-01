"""Regression tests for the API-to-MCP integration gate."""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "verify_api_integration", ROOT / "scripts" / "verify_api_integration.py"
)
assert SPEC and SPEC.loader
verify = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(verify)


def test_cli_help_gate_cannot_rewrite_the_lockfile() -> None:
    config = yaml.safe_load(
        (ROOT / ".pre-commit-config.yaml").read_text(encoding="utf-8")
    )
    hooks = [hook for repo in config["repos"] for hook in repo.get("hooks", [])]
    cli_help = next(hook for hook in hooks if hook["id"] == "check-cli-help")
    assert "uv run --frozen python" in cli_help["entry"]


def test_microsoft_action_catalogs_meet_the_existing_baseline() -> None:
    """The current routers map every public API method except lifecycle close."""
    result = verify.verify_agent(ROOT)
    assert result is not None
    assert result["agent_name"] == "microsoft-agent"
    assert result["total_methods"] == 269
    assert result["covered_methods"] == 268
    assert result["coverage"] == pytest.approx(99.628, abs=0.001)
    assert result["unmapped"] == ["close"]


def test_only_catalogs_consumed_by_a_tool_count_as_integration(tmp_path: Path) -> None:
    mcp_source = tmp_path / "mcp_domain.py"
    mcp_source.write_text(
        '_USED_ACTIONS = ("list_items",)\n'
        '_STALE_ACTIONS = ("delete_item",)\n\n'
        "def register_domain_tools(mcp):\n"
        "    @mcp.tool()\n"
        "    async def demo_domain(action, client):\n"
        "        action = resolve_action(action, _USED_ACTIONS)\n"
        "        return await getattr(client, action)()\n\n"
        "    @mcp.tool()\n"
        "    async def demo_policy_only(action, client):\n"
        "        resolve_action(action, _STALE_ACTIONS)\n"
        "        return {'allowed': True}\n",
        encoding="utf-8",
    )
    api_methods: dict[str, dict[str, object]] = {
        "list_items": {},
        "delete_item": {},
    }
    _, mapped = verify.parse_mcp_server([mcp_source], api_methods)
    assert mapped == {"list_items"}


def test_dot_prefixed_worktree_path_is_not_mistaken_for_hidden_source(
    tmp_path: Path,
) -> None:
    root = tmp_path / ".state" / "demo-agent"
    package = root / "demo_agent"
    mcp_package = package / "mcp"
    mcp_package.mkdir(parents=True)
    (package / "api_client.py").write_text(
        "class DemoApiClient:\n    async def list_items(self):\n        return []\n",
        encoding="utf-8",
    )
    (package / "mcp_server.py").write_text(
        "from demo_agent.mcp.domain import register_domain_tools\n",
        encoding="utf-8",
    )
    (mcp_package / "domain.py").write_text(
        '_DOMAIN_ACTIONS = ("list_items",)\n\n'
        "def register_domain_tools(mcp):\n"
        "    @mcp.tool()\n"
        "    async def demo_domain(action, client):\n"
        "        action = resolve_action(action, _DOMAIN_ACTIONS)\n"
        "        return await getattr(client, action)()\n",
        encoding="utf-8",
    )
    result = verify.verify_agent(root)
    assert result is not None
    assert result["covered_methods"] == result["total_methods"] == 1
