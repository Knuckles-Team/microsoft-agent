"""Action-discovery tests for the current mixed MCP tool surface.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

import sys
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from microsoft_agent.mcp_server import get_mcp_instance

_ACTION_ENVELOPE_FIELDS = {"action", "params_json"}


def _action_routed_tools(tools):
    """Return only tools exposing the current action-routing envelope."""

    return [
        tool
        for tool in tools
        if set(tool.parameters.get("properties", {})) == _ACTION_ENVELOPE_FIELDS
    ]


@pytest.mark.concept("AU-ECO.mcp.fastmcp-middleware")
@pytest.mark.asyncio
async def test_list_actions_returns_action_names():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    tools = await mcp.list_tools()
    action_tools = _action_routed_tools(tools)
    assert action_tools, "expected registered action-routed tools"
    assert len(action_tools) < len(tools), "expected a mixed MCP tool surface"

    client = MagicMock()
    ctx = AsyncMock()

    for tool in action_tools:
        result = await tool.fn(
            action="list_actions",
            params_json="{}",
            client=client,
            ctx=ctx,
        )
        assert isinstance(result, dict)
        assert result["service"] == "microsoft-agent"
        assert isinstance(result["actions"], list)
        assert result["actions"], f"{tool.fn.__name__} returned no actions"


@pytest.mark.concept("AU-ECO.mcp.fastmcp-middleware")
@pytest.mark.asyncio
async def test_unknown_action_raises_with_discovery_hint():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    action_tools = _action_routed_tools(await mcp.list_tools())
    assert action_tools, "expected registered action-routed tools"

    with pytest.raises(ValueError, match="list_actions"):
        await action_tools[0].fn(
            action="definitely_not_a_real_action",
            params_json="{}",
            client=MagicMock(),
            ctx=AsyncMock(),
        )
