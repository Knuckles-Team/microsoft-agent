import pytest


@pytest.mark.concept("ECO-4.1")
def test_server_startup(monkeypatch):
    """Validates that the server module can start successfully."""
    import sys
    from unittest.mock import MagicMock

    from microsoft_agent import agent_server as server_module

    monkeypatch.setattr(
        sys,
        "argv",
        ["agent_server.py", "--mcp-url", "http://localhost:8000", "--debug"],
    )

    mock_create_agent_server = MagicMock()
    mock_initialize_workspace = MagicMock()
    mock_load_identity = MagicMock(
        return_value={"name": "Microsoft Agent", "description": "AI agent"}
    )

    monkeypatch.setattr(server_module, "create_agent_server", mock_create_agent_server)
    monkeypatch.setattr(server_module, "initialize_workspace", mock_initialize_workspace)
    monkeypatch.setattr(server_module, "load_identity", mock_load_identity)

    server_module.agent_server()

    assert mock_create_agent_server.called
    print("Startup tests handled correctly.")
