from typing import Any

from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiSystem(MicrosoftGraphApiBase):
    def login(self, force: bool = False) -> str:
        """Authenticate with Microsoft."""
        if not force:
            token = self.auth_manager.get_token()
            if token:
                return "Already authenticated."

        return self.auth_manager.login()

    def logout(self) -> str:
        """Logout."""
        self.auth_manager.logout()
        return "Logged out."

    def verify_login(self) -> str:
        """Verify login status."""
        token = self.auth_manager.get_token()
        if token:
            account = self.auth_manager.get_current_account()
            if account:
                return f"Authenticated as {account.get('username', 'Unknown')}"
            return "Authenticated with workload identity"
        return "Not authenticated."

    def list_accounts(self) -> list[dict[str, Any]]:
        """List accounts."""
        return self.auth_manager.list_accounts()

    def search_tools(self, query: str, limit: int = 10) -> list[str]:
        """Search methods in this class."""

        matches = []
        for name in dir(self):
            if name.startswith("_"):
                continue
            if query.lower() in name.lower():
                matches.append(name)
            if len(matches) >= limit:
                break
        return matches
