"""Tests for the typed Microsoft MCP surface: one tool per Graph operation."""

from __future__ import annotations

import importlib

import pytest
from _served import served_tools as _served_tools
from fastmcp import FastMCP

from microsoft_agent import mcp_server


@pytest.mark.spec("CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT-R001")
def test_served_tools_are_typed_operations_without_action_routing() -> None:
    tools = _served_tools()
    names = {tool.name for tool in tools}

    assert len(names) > 300
    assert {
        "microsoft_list_mail_messages",
        "microsoft_send_mail",
        "microsoft_submit_print_document",
        "microsoft_list_users",
        "microsoft_format_excel_range",
    } <= names
    for tool in tools:
        properties = (tool.inputSchema or {}).get("properties", {})
        assert "action" not in properties, tool.name
        assert tool.description, tool.name


@pytest.mark.spec("CONN-MICROSOFT-AGENT-CONNECTOR-CONTENT-R001")
def test_old_condensed_family_tools_are_gone() -> None:
    names = {tool.name for tool in _served_tools()}

    assert not names & {"microsoft_mail", "microsoft_files", "microsoft_auth"}


@pytest.mark.asyncio
async def test_knowledge_projection_registers_only_current_tools() -> None:
    server = FastMCP("test-microsoft-agent")

    mcp_server.register_kg_tools(server)

    tools = await server.list_tools()
    assert {tool.name for tool in tools} == {"list_microsoft_ingestion_projection"}
    projection = next(
        tool for tool in tools if tool.name == "list_microsoft_ingestion_projection"
    )
    assert {"kg", "read"}.issubset(projection.tags)


def test_fragmented_registration_and_auth_surfaces_are_absent() -> None:
    module = importlib.import_module("microsoft_agent.mcp_server")

    assert not hasattr(module, "register_prompts")
    assert not hasattr(module, "register_misc_tools")
    assert not hasattr(module, "AuthManager")


def test_mcp_server_version_is_explicit() -> None:
    assert isinstance(mcp_server.__version__, str)
    assert mcp_server.__version__
