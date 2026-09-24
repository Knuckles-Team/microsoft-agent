from abc import ABC, abstractmethod
from urllib.parse import urlparse

import httpx
from agent_connector_sdk.exceptions import AuthError
from agent_connector_sdk.http.client import create_async_http_client
from agent_connector_sdk.http.options import HttpClientOptions
from agent_connector_sdk.tls.profile import ResolvedTLSProfile
from agent_connector_sdk.tls.resolve import resolve_tls_profile
from kiota_authentication_azure.azure_identity_authentication_provider import (
    AzureIdentityAuthenticationProvider,
)
from msgraph import GraphServiceClient
from msgraph.graph_request_adapter import GraphRequestAdapter

from microsoft_agent.auth import AuthManager
from microsoft_agent.credential_adapter import AuthManagerCredential


class MicrosoftGraphApiBase(ABC):
    @abstractmethod
    def verify_login(self) -> str:
        """Return the authenticated status for the composed client."""

    def __init__(self, auth_manager: AuthManager):
        self.auth_manager = auth_manager
        self.credential = AuthManagerCredential(auth_manager)
        status = self.verify_login()
        if "Not authenticated" in status:
            raise AuthError(f"Microsoft authentication failed: {status}")

        endpoint_host = urlparse(auth_manager.graph_base_url).hostname
        if endpoint_host is None:
            raise ValueError("Microsoft Graph endpoint is invalid")
        self.tls_profile: ResolvedTLSProfile | None = resolve_tls_profile(
            "microsoft_graph",
            profile_name=auth_manager.graph_tls_profile,
            profile_ref=auth_manager.graph_tls_profile_ref,
        )
        try:
            authentication_provider = AzureIdentityAuthenticationProvider(
                self.credential,
                scopes=auth_manager.scopes,
                allowed_hosts=[endpoint_host],
            )
            # SDK-GAP: agent_connector_sdk.http has no pin_egress/allowed_private_hosts
            # equivalent (AU's SSRF-pinning knobs on create_async_http_client) -- see
            # SDK-GAPS.md #12. follow_redirects=False is now the governed client's
            # fixed, non-configurable behavior (per its own docstring), so it is no
            # longer passed explicitly.
            self._http_client = create_async_http_client(
                HttpClientOptions(
                    base_url=auth_manager.graph_base_url,
                    timeout=httpx.Timeout(30.0),
                    tls=self.tls_profile,
                    limits=httpx.Limits(
                        max_connections=64,
                        max_keepalive_connections=16,
                    ),
                )
            )
        except Exception:
            self.tls_profile.cleanup()
            self.tls_profile = None
            raise
        request_adapter = GraphRequestAdapter(
            authentication_provider,
            client=self._http_client,
        )
        request_adapter.base_url = auth_manager.graph_base_url
        self.client = GraphServiceClient(request_adapter=request_adapter)

    async def close(self) -> None:
        """Release the provider transport and materialized TLS profile."""

        profile, self.tls_profile = self.tls_profile, None
        if profile is None:
            return
        try:
            await self._http_client.aclose()
        finally:
            profile.cleanup()
