"""Audience-aware token acquisition for Microsoft integration clients."""

from __future__ import annotations

import asyncio
from collections.abc import Iterable
from urllib.parse import urlparse

from microsoft_agent.auth import (
    AuthenticationRequiredError,
    AuthManager,
    get_auth_manager,
)


def normalize_audience(audience: str) -> str:
    """Validate and normalize an OAuth resource audience."""

    value = audience.strip().rstrip("/")
    parsed = urlparse(value)
    if (
        parsed.scheme not in {"https", "api"}
        or not parsed.netloc
        or parsed.username
        or parsed.password
        or parsed.query
        or parsed.fragment
    ):
        raise ValueError(
            "OAuth audiences must be absolute HTTPS or api:// resource URLs"
        )
    return value


class MicrosoftAudienceTokenProvider:
    """Acquire tokens for an explicit set of Microsoft resource audiences.

    MSAL is synchronous, so acquisition runs in a worker thread.  The provider
    accepts only configured audiences; an invocation cannot choose a new token
    target dynamically.
    """

    def __init__(
        self,
        allowed_audiences: Iterable[str],
        auth_manager: AuthManager | None = None,
        *,
        allow_interactive: bool = False,
    ) -> None:
        audiences = frozenset(normalize_audience(item) for item in allowed_audiences)
        if not audiences:
            raise ValueError("at least one OAuth audience must be allowlisted")
        self._allowed_audiences = audiences
        self._auth_manager = auth_manager
        self._allow_interactive = allow_interactive

    @property
    def allowed_audiences(self) -> frozenset[str]:
        """Return the immutable normalized resource allowlist."""

        return self._allowed_audiences

    async def get_token(self, audience: str) -> str:
        """Return an access token for one allowlisted resource audience."""

        normalized = normalize_audience(audience)
        if normalized not in self._allowed_audiences:
            raise AuthenticationRequiredError(
                "Token acquisition was denied for an unconfigured resource audience."
            )
        manager = self._auth_manager or get_auth_manager()
        scope = f"{normalized}/.default"
        token = await asyncio.to_thread(
            manager.acquire_token_for_scopes,
            [scope],
            allow_interactive=self._allow_interactive,
        )
        if not token:
            raise AuthenticationRequiredError(
                "Microsoft authentication is required for this integration."
            )
        return token
