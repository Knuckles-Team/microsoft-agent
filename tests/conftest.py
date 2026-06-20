"""Global pytest configuration and fixtures.

CONCEPT:ECO-4.1
"""

import atexit

# Globally disable atexit hook registration during tests to prevent real keyring calls during exit
atexit.register = lambda fn, *args, **kwargs: fn

from unittest.mock import patch

import pytest


@pytest.fixture(autouse=True, scope="session")
def mock_system_keyring():
    """Globally mock keyring to prevent database/SecretService keyring hangs in headless CI/test environments."""
    with (
        patch("keyring.get_password", return_value=None),
        patch("keyring.set_password", return_value=None),
        patch("keyring.delete_password", return_value=None),
    ):
        yield
