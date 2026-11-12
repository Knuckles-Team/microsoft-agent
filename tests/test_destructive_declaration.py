"""Destructive operations are declared with the SDK and ask before they run."""

from __future__ import annotations

import pytest

from microsoft_agent.mcp_server import get_mcp_instance
from microsoft_agent.tool_policy import ToolRisk, classify_tool_risk

DECLARED_GRAPH = (
    "add_application_password create_application create_conditional_access_policy "
    "create_role_assignment create_service_principal dismiss_risky_user "
    "retire_managed_device update_admin_sharepoint update_conditional_access_policy "
    "update_organization wipe_managed_device"
).split()
HAND_WRITTEN = (
    "delete_powerpoint_slide_in_office",
    "cancel_power_automate_desktop_flow_run",
    "reboot_intune_device",
    "shut_down_intune_device",
    "submit_windows_action",
)


async def _tools(monkeypatch):
    monkeypatch.setenv("MICROSOFT_ALLOW_WRITES", "true")
    monkeypatch.setattr("sys.argv", ["microsoft-mcp"])
    server = [x for x in get_mcp_instance() if hasattr(x, "list_tools")][0]
    return server, {tool.name: tool for tool in await server.list_tools()}


@pytest.mark.asyncio
async def test_destructive_tools_are_declared_and_take_confirm(monkeypatch):
    _, tools = await _tools(monkeypatch)
    names = [f"microsoft_{name}" for name in DECLARED_GRAPH] + [
        name
        for name in tools
        if classify_tool_risk(name) is ToolRisk.DESTRUCTIVE
        and name.removeprefix("microsoft_").split("_", 1)[0]
        in {"delete", "remove", "wipe", "retire", "cancel"}
    ]
    names += [name for name in HAND_WRITTEN if name in tools]
    assert names
    for name in names:
        assert name in tools, name
        properties = tools[name].parameters["properties"]
        assert tools[name].annotations.destructive_hint is True, name
        assert "confirm" in properties, name
        assert not [key for key in properties if key.startswith("allow_")], name


def test_connector_defines_no_destructive_setting():
    from microsoft_agent.settings import MicrosoftSettings

    assert not hasattr(MicrosoftSettings(), "allow_destructive_tools")
    assert "destructive_action_acknowledged" not in _evidence_fields()


def _evidence_fields():
    from microsoft_agent.intune_service import ConfirmationEvidence

    return ConfirmationEvidence.model_fields


@pytest.mark.asyncio
async def test_hand_written_tool_refuses_without_confirmation(monkeypatch):
    server, tools = await _tools(monkeypatch)
    called = []

    async def refuse(ctx, operation, arguments=None, *, confirm=False):
        called.append((operation, confirm))
        return {"cancelled": True, "operation": operation}

    monkeypatch.setattr("microsoft_agent.office_bridge.guard_destructive", refuse)
    tool = await server.get_tool("delete_powerpoint_slide_in_office")
    result = await tool.fn(
        session_id="00000000-0000-4000-8000-000000000000", slide_number=1
    )
    assert result == {
        "cancelled": True,
        "operation": "delete_powerpoint_slide_in_office",
    }
    assert called == [("delete_powerpoint_slide_in_office", False)]
