"""Regression test for BUG-CX-046.

Every ``mcp_*.py`` tool module in ``microsoft_agent/mcp/`` contained the pattern::

    if ctx:
        ctx.info("Executing tool...")

``Context.info`` (fastmcp) is an async coroutine method. Calling it without
``await`` creates a coroutine object that is immediately discarded -- the log
line never emits and Python raises ``RuntimeWarning: coroutine 'Context.info'
was never awaited``.

This test exercises one representative tool (``microsoft_auth``) with a mock
``Context`` whose ``.info`` is an ``AsyncMock`` and asserts it was actually
awaited. Before the fix, ``ctx.info(...)`` is called-but-not-awaited, so
``AsyncMock.assert_awaited_once()`` fails because the mock's await count stays
0 even though it was *called*. After the fix (``await ctx.info(...)``), the
mock is awaited exactly once.
"""

from __future__ import annotations

from unittest.mock import AsyncMock

import pytest
from fastmcp import FastMCP

from microsoft_agent.mcp_server import register_auth_tools


@pytest.mark.asyncio
async def test_microsoft_auth_tool_awaits_ctx_info() -> None:
    mcp = FastMCP("bug-cx-046-test")
    register_auth_tools(mcp)
    tool = await mcp.get_tool("microsoft_auth")

    client = AsyncMock()
    client.list_accounts = AsyncMock(return_value=[])
    mock_ctx = AsyncMock()

    await tool.fn(
        action="list_accounts",
        params_json="{}",
        client=client,
        ctx=mock_ctx,
    )

    mock_ctx.info.assert_awaited_once()
