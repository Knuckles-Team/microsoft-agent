import runpy
from unittest.mock import patch

import pytest


@pytest.mark.concept("ECO-4.1")
def test_main_py():
    with patch("microsoft_agent.agent_server.agent_server") as mock_agent_server:
        runpy.run_module("microsoft_agent.__main__", run_name="__main__")
        mock_agent_server.assert_called_once()
