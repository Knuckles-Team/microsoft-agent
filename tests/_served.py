"""Shared helper: the tools the declared connector serves."""

from __future__ import annotations


def served_tools() -> list:
    """Every tool the declared connector serves, as the platform builds it."""
    import asyncio

    from fastmcp import Client
    from scion_sdk.connector import build_server

    mcp = build_server("microsoft_agent", command_args=[])[1]

    async def read() -> list:
        async with Client(mcp) as client:
            return await client.list_tools()

    return asyncio.run(read())
