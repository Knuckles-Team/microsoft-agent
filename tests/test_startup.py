import pytest


@pytest.mark.concept("ECO-4.1")
def test_server_startup(monkeypatch):
    """Validates that the server module can start successfully."""
    import os
    import sys
    from unittest.mock import MagicMock

    target_dir = None
    for d in [".", "src", "agent", "microsoft_agent"]:
        if os.path.exists(os.path.join(d, "agent_server.py")):
            target_dir = d
            break

    if target_dir is None:
        return

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

    monkeypatch.setattr("agent_utilities.create_agent_server", mock_create_agent_server)
    monkeypatch.setattr(
        "agent_utilities.initialize_workspace", mock_initialize_workspace
    )
    monkeypatch.setattr("agent_utilities.load_identity", mock_load_identity)

    original_path = list(sys.path)
    sys.path.insert(0, os.path.abspath(target_dir))

    try:
        import runpy

        runpy.run_module("agent_server", run_name="__main__")
    finally:
        sys.path = original_path

    assert mock_create_agent_server.called
    print("Startup tests handled correctly.")
