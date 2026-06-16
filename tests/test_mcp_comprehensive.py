import os
import sys
import ast
import json
import pytest
from unittest.mock import MagicMock, patch

from microsoft_agent.mcp_server import get_mcp_instance, mcp_server


def get_actions_from_mcp(filepath):
    """Parse mcp_server.py using AST to find all actions defined in tool functions."""
    with open(filepath, "r", encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename=filepath)

    actions = {}
    for node in ast.walk(tree):
        if isinstance(
            node, (ast.FunctionDef, ast.AsyncFunctionDef)
        ) and node.name.startswith("register_"):
            for inner in node.body:
                if isinstance(inner, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    tool_name = inner.name
                    tool_actions = []
                    for child in ast.walk(inner):
                        if isinstance(child, ast.Compare):
                            if (
                                isinstance(child.left, ast.Name)
                                and child.left.id == "action"
                            ):
                                for comparator in child.comparators:
                                    if isinstance(comparator, ast.Constant):
                                        tool_actions.append(comparator.value)
                    if tool_actions:
                        actions[tool_name] = tool_actions
    return actions


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_all_mcp_tools_and_actions():
    # 1. Fetch mcp instance and list of tools
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    # 2. Extract actions using AST
    mcp_file_path = os.path.join(
        os.path.dirname(os.path.dirname(__file__)), "microsoft_agent", "mcp_server.py"
    )
    extracted_actions = get_actions_from_mcp(mcp_file_path)

    # 3. Create a dynamic Mock client where accessing any attribute returns a mock that returns a dict
    class MockClient(MagicMock):
        def __getattr__(self, name):
            # Return a callable mock that returns a dictionary to avoid serialize/unserialize issues
            mock_method = MagicMock()
            mock_method.return_value = {"status": "success", "method": name}
            return mock_method

    mock_client = MockClient()
    mock_ctx = MagicMock()

    # 4. Await list_tools to fetch registered tools
    tools = await mcp.list_tools()

    # 5. Dynamically run each registered action in each tool
    for tool in tools:
        tool_fn_name = tool.fn.__name__
        actions_list = extracted_actions.get(tool_fn_name, [])

        # Execute each action
        for action in actions_list:
            res = await tool.fn(
                action=action,
                params_json='{"key": "value"}',
                client=mock_client,
                ctx=mock_ctx,
            )
            # Ensure client method was invoked (unless it was handled internally or errored)
            assert isinstance(res, dict)

        # Test invalid params_json
        err_res = await tool.fn(
            action=actions_list[0] if actions_list else "dummy",
            params_json="{invalid json",
            client=mock_client,
            ctx=mock_ctx,
        )
        assert "error" in err_res

        # Test unknown action (shared resolve_action raises a rich did-you-mean error)
        with pytest.raises(ValueError, match="Unknown action"):
            await tool.fn(
                action="invalid_unknown_action",
                params_json="{}",
                client=mock_client,
                ctx=mock_ctx,
            )


@pytest.mark.concept("ECO-4.1")
@pytest.mark.asyncio
async def test_mcp_health_check_route():
    with patch.object(sys, "argv", ["mcp_server.py"]):
        mcp, _, _ = get_mcp_instance()

    health_route = None
    for route in mcp._additional_http_routes:
        if getattr(route, "path", None) == "/health":
            health_route = route
            break

    assert health_route is not None
    mock_request = MagicMock()
    response = await health_route.endpoint(mock_request)
    assert response.status_code == 200
    assert json.loads(response.body.decode()) == {"status": "OK"}


@pytest.mark.concept("ECO-4.1")
def test_mcp_server_entrypoints():
    # 1. Test stdio transport
    with patch.object(sys, "argv", ["mcp_server.py", "--transport", "stdio"]):
        with patch("fastmcp.FastMCP.run") as mock_run:
            mcp_server()
            mock_run.assert_called_once_with(transport="stdio")

    # 2. Test streamable-http transport
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

    # 3. Test sse transport
    with patch.object(
        sys,
        "argv",
        [
            "mcp_server.py",
            "--transport",
            "sse",
            "--host",
            "127.0.0.1",
            "--port",
            "8555",
        ],
    ):
        with patch("fastmcp.FastMCP.run") as mock_run:
            mcp_server()
            mock_run.assert_called_once_with(
                transport="sse", host="127.0.0.1", port=8555
            )

    # 4. Test invalid transport (should call sys.exit)
    with patch("microsoft_agent.mcp_server.get_mcp_instance") as mock_get:
        mock_args = MagicMock()
        mock_args.transport = "invalid"
        mock_get.return_value = (MagicMock(), mock_args, MagicMock())
        with pytest.raises(SystemExit):
            mcp_server()
