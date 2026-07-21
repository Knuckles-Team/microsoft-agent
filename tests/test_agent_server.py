import sys
from unittest.mock import patch

import pytest

from microsoft_agent.agent_server import agent_server


@pytest.mark.concept("ECO-4.1")
def test_agent_server_cli():
    with patch(
        "microsoft_agent.agent_server.create_agent_server"
    ) as mock_create_server:
        with (
            patch("microsoft_agent.agent_server.initialize_workspace"),
            patch(
                "microsoft_agent.agent_server.load_identity",
                return_value={"name": "Microsoft Agent", "content": "Prompt"},
            ),
            patch(
                "microsoft_agent.agent_server.setting",
                side_effect=lambda _key, default: default,
            ),
            patch.object(
                sys,
                "argv",
                [
                    "agent_server.py",
                    "--mcp-url",
                    "http://localhost:8000",
                    "--host",
                    "127.0.0.1",
                    "--port",
                    "8001",
                    "--debug",
                ],
            ),
        ):
            agent_server()
            mock_create_server.assert_called_once()
            kwargs = mock_create_server.call_args[1]
            assert kwargs["mcp_url"] == "http://localhost:8000"
            assert kwargs["host"] == "127.0.0.1"
            assert kwargs["port"] == 8001
            assert kwargs["debug"] is True
            assert kwargs["name"] == "Microsoft Agent"
            assert kwargs["system_prompt"] == "Prompt"
