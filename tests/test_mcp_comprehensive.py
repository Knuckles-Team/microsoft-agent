import json
import sys
from unittest.mock import MagicMock, patch

import pytest

from microsoft_agent.mcp_server import get_mcp_instance, mcp_server


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
        # The served-security profile (SDK >=0.3.6) refuses a network transport
        # without an auth provider; this test covers argument routing only.
        with (
            patch("fastmcp.FastMCP.run") as mock_run,
            patch("scion_sdk.mcp.server.apply_served_security_profile"),
        ):
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
