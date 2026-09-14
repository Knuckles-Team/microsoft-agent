"""Regression tests for generated Agent Utilities imports."""

from __future__ import annotations

import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "generate_code", ROOT / "scripts" / "generate_code.py"
)
assert SPEC and SPEC.loader
generate_code = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(generate_code)


def test_generated_mcp_uses_canonical_agent_utilities_import() -> None:
    generated = generate_code.generate_mcp_code([])

    assert "from agent_utilities import to_boolean" in generated
    assert "agent_utilities.agent_utilities" not in generated


def test_generated_agent_uses_canonical_agent_utilities_import() -> None:
    generated = generate_code.generate_agent_code([])

    assert "from agent_utilities import create_model" in generated
    assert "agent_utilities.agent_utilities" not in generated
