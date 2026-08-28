import json
import sys
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from microsoft_agent.mcp_server import get_mcp_instance, mcp_server

_ACTION_ENVELOPE_FIELDS = {"action", "params_json"}
_ACTION_TOOL_NAMES = {
    "microsoft_admin",
    "microsoft_agreements",
    "microsoft_applications",
    "microsoft_audit",
    "microsoft_auth",
    "microsoft_calendar",
    "microsoft_chat",
    "microsoft_communications",
    "microsoft_connections",
    "microsoft_contacts",
    "microsoft_devices",
    "microsoft_directory",
    "microsoft_domains",
    "microsoft_education",
    "microsoft_employee_experience",
    "microsoft_files",
    "microsoft_groups",
    "microsoft_identity",
    "microsoft_mail",
    "microsoft_meta",
    "microsoft_notes",
    "microsoft_organization",
    "microsoft_places",
    "microsoft_policies",
    "microsoft_print",
    "microsoft_privacy",
    "microsoft_reports",
    "microsoft_search",
    "microsoft_security",
    "microsoft_sites",
    "microsoft_solutions",
    "microsoft_storage",
    "microsoft_subscriptions",
    "microsoft_tasks",
    "microsoft_teams",
    "microsoft_user",
}
_CORE_DEDICATED_SCHEMAS = {
    "list_microsoft_ingestion_projection": {
        "drive_id",
        "drive_item_id",
        "kind",
        "params_json",
    },
}


def _tool_properties(tool) -> dict:
    schema = tool.parameters
    assert isinstance(schema, dict), f"{tool.name} must publish a JSON schema"
    assert schema.get("type") == "object", f"{tool.name} schema must be an object"
    properties = schema.get("properties")
    assert isinstance(properties, dict), f"{tool.name} must publish properties"
    assert set(schema.get("required", ())) <= set(properties)
    assert all(isinstance(value, dict) for value in properties.values())
    return properties


def _is_action_routed(tool) -> bool:
    return set(_tool_properties(tool)) == _ACTION_ENVELOPE_FIELDS


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_all_mcp_tools_and_actions():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    tools = await mcp.list_tools()
    action_tools = [tool for tool in tools if _is_action_routed(tool)]
    dedicated_tools = [tool for tool in tools if not _is_action_routed(tool)]

    assert {tool.name for tool in action_tools} == _ACTION_TOOL_NAMES
    assert dedicated_tools, "expected dedicated tools alongside action-routed tools"

    mock_client = MagicMock()
    mock_ctx = AsyncMock()
    for tool in action_tools:
        discovery = await tool.fn(
            action="list_actions",
            params_json="{}",
            client=mock_client,
            ctx=mock_ctx,
        )
        assert discovery["service"] == "microsoft-agent"
        assert discovery["actions"]

        invalid_json = await tool.fn(
            action=discovery["actions"][0],
            params_json="{invalid json",
            client=mock_client,
            ctx=mock_ctx,
        )
        assert invalid_json == {"error": "Invalid params_json"}

        with pytest.raises(ValueError, match="Unknown action"):
            await tool.fn(
                action="invalid_unknown_action",
                params_json="{}",
                client=mock_client,
                ctx=mock_ctx,
            )

    dedicated_by_name = {tool.name: tool for tool in dedicated_tools}
    assert _CORE_DEDICATED_SCHEMAS.keys() <= dedicated_by_name.keys()
    for name, expected_properties in _CORE_DEDICATED_SCHEMAS.items():
        assert set(_tool_properties(dedicated_by_name[name])) == expected_properties

    for tool in dedicated_tools:
        properties = _tool_properties(tool)
        assert set(properties) != _ACTION_ENVELOPE_FIELDS
        assert not _ACTION_ENVELOPE_FIELDS <= set(properties)


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_mcp_health_check_route():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    health_route = None
    for route in mcp._additional_http_routes:
        if (
            getattr(route, "path", None) == "/health"
            and route.endpoint.__module__ == "microsoft_agent.mcp_server"
        ):
            health_route = route
            break

    assert health_route is not None
    mock_request = MagicMock()
    response = await health_route.endpoint(mock_request)
    assert response.status_code == 200
    assert json.loads(response.body.decode()) == {"status": "OK"}


@pytest.mark.concept("ECO-4.1")
def test_mcp_server_entrypoints():
    with patch.object(sys, "argv", ["mcp_server.py", "--transport", "stdio"]):
        with patch("fastmcp.FastMCP.run") as mock_run:
            mcp_server()
            mock_run.assert_called_once_with(transport="stdio")

    with patch.object(
        sys,
        "argv",
        [
            "mcp_server.py",
            "--transport",
            "streamable-http",
            "--host",
            "127.0.0.1",
            "--port",
            "8555",
        ],
    ):
        with patch("fastmcp.FastMCP.run") as mock_run:
            mcp_server()
            mock_run.assert_called_once_with(
                transport="streamable-http", host="127.0.0.1", port=8555
            )

    with patch("microsoft_agent.mcp_server.get_mcp_instance") as mock_get:
        mock_args = MagicMock()
        mock_args.transport = "invalid"
        mock_get.return_value = (MagicMock(), mock_args, MagicMock())
        with pytest.raises(SystemExit):
            mcp_server()
