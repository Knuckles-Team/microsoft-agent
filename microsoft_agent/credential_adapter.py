from typing import Any, Optional
from azure.core.credentials import AccessToken, TokenCredential
import time

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
        claims: Optional[str] = None,
        tenant_id: Optional[str] = None,
        **kwargs: Any,
    ) -> AccessToken:
        # Note: We ignore the requested scopes here and used the ones configured in AuthManager
        # because MSAL's acquire_token_silent usually needs the original scopes or a subset.
        # The SDK usually requests 'https://graph.microsoft.com/.default', which might work
        # if the app is configured right, but for now we stick to what we have.

        token_details = self.auth_manager.get_token_details()

        if not token_details:
            # If we can't get a token silently, we might need to trigger a login flow
            # But this method is usually expected to be non-interactive or raise.
            # Since our agent is interactive via tools, we might just raise
            # and let the user use the 'login' tool if needed.
            pass
            # Try getting a token even if it might fail, to trigger error

        if token_details and "access_token" in token_details:
            # MSAL returns expires_in, we need expires_on (timestamp)
            # It usually also returns "expires_on" in the result if it's from cache?
            # Let's check.
            # If not, we calculate it.
            expires_on = token_details.get("expires_on")
            if not expires_on:
                expires_in = token_details.get("expires_in", 3600)
                expires_on = int(time.time()) + int(expires_in)

            return AccessToken(token_details["access_token"], int(expires_on))

        raise Exception(
            "Failed to acquire token. Please use the 'login' tool to authenticate."
        )
