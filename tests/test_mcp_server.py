"""Current-only tests for the condensed Microsoft MCP surface."""

from __future__ import annotations

import importlib

import pytest
from fastmcp import FastMCP

from microsoft_agent import mcp_server

_ACTION_FAMILIES = (
    ("auth", "microsoft_auth"),
    ("meta", "microsoft_meta"),
    ("mail", "microsoft_mail"),
    ("files", "microsoft_files"),
    ("calendar", "microsoft_calendar"),
    ("notes", "microsoft_notes"),
    ("tasks", "microsoft_tasks"),
    ("contacts", "microsoft_contacts"),
    ("user", "microsoft_user"),
    ("chat", "microsoft_chat"),
    ("teams", "microsoft_teams"),
    ("sites", "microsoft_sites"),
    ("search", "microsoft_search"),
    ("groups", "microsoft_groups"),
    ("admin", "microsoft_admin"),
    ("organization", "microsoft_organization"),
    ("domains", "microsoft_domains"),
    ("subscriptions", "microsoft_subscriptions"),
    ("communications", "microsoft_communications"),
    ("identity", "microsoft_identity"),
    ("security", "microsoft_security"),
    ("audit", "microsoft_audit"),
    ("reports", "microsoft_reports"),
    ("applications", "microsoft_applications"),
    ("directory", "microsoft_directory"),
    ("policies", "microsoft_policies"),
    ("devices", "microsoft_devices"),
    ("education", "microsoft_education"),
    ("agreements", "microsoft_agreements"),
    ("places", "microsoft_places"),
    ("print", "microsoft_print"),
    ("privacy", "microsoft_privacy"),
    ("solutions", "microsoft_solutions"),
    ("storage", "microsoft_storage"),
    ("employee_experience", "microsoft_employee_experience"),
    ("connections", "microsoft_connections"),
)


@pytest.mark.parametrize(("family", "tool_name"), _ACTION_FAMILIES)
@pytest.mark.asyncio
async def test_each_current_action_family_registers_one_described_tool(
    family: str,
    tool_name: str,
) -> None:
    server = FastMCP("test-microsoft-agent")
    register = getattr(mcp_server, f"register_{family}_tools")

    register(server)

    tools = await server.list_tools()
    assert [tool.name for tool in tools] == [tool_name]
    assert tools[0].description
    assert family in tools[0].tags


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
