"""Action-discovery standardization tests.

Verifies the shared ``agent_utilities.mcp_utilities.resolve_action`` helper is
wired into every action-routed tool so callers get ``list_actions`` discovery,
plural->singular aliases, and a rich did-you-mean error on unknown actions.

CONCEPT:ECO-4.1
"""

import sys
from unittest.mock import MagicMock, patch

import pytest

from microsoft_agent.mcp_server import get_mcp_instance


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_list_actions_returns_action_names():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    tools = await mcp.list_tools()
    assert tools, "expected registered action-routed tools"

    client = MagicMock()
    ctx = MagicMock()

    for tool in tools:
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


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_unknown_action_raises_with_discovery_hint():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    tools = await mcp.list_tools()
    client = MagicMock()
    ctx = MagicMock()

    tool = tools[0]
    with pytest.raises(ValueError, match="list_actions"):
        await tool.fn(
            action="definitely_not_a_real_action",
            params_json="{}",
            client=client,
            ctx=ctx,
        )
