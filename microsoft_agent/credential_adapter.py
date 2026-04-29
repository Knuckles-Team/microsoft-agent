import time
from typing import Any

from azure.core.credentials import AccessToken, TokenCredential

from microsoft_agent.auth import AuthManager


class AuthManagerCredential(TokenCredential):
    """
    Adapter to allow AuthManager to be used as a TokenCredential
    for use with the Microsoft Graph SDK.
    """

    def __init__(self, auth_manager: AuthManager):
        self.auth_manager = auth_manager

    def get_token(
        self,
        *scopes: str,
        claims: str | None = None,
        tenant_id: str | None = None,
        **kwargs: Any,
    ) -> AccessToken:

        token_details = self.auth_manager.get_token_details()

        if not token_details:

            pass

        if token_details and "access_token" in token_details:

            expires_on = token_details.get("expires_on")
            if not expires_on:
                expires_in = token_details.get("expires_in", 3600)
                expires_on = int(time.time()) + int(expires_in)

            return AccessToken(token_details["access_token"], int(expires_on))

        raise Exception(
            "Failed to acquire token. Please use the 'login' tool to authenticate."
        )
