"""Fail-closed MCP tool policy for Microsoft Agent side effects."""

from __future__ import annotations

import fnmatch
from collections.abc import Mapping
from enum import StrEnum
from typing import Any

from fastmcp.exceptions import ToolError
from fastmcp.server.middleware import Middleware, MiddlewareContext
from pydantic import BaseModel, ConfigDict

from microsoft_agent.settings import MicrosoftSettings, get_settings


class ToolRisk(StrEnum):
    """Side-effect classification used by the runtime authorization gate."""

    READ = "read"
    WRITE = "write"
    DESTRUCTIVE = "destructive"


READ_PREFIXES = (
    "get_",
    "list_",
    "search_",
    "find_",
    "download_",
    "verify_",
    "check_",
)
WRITE_PREFIXES = (
    "add_",
    "create_",
    "format_",
    "invite_",
    "move_",
    "reply_",
    "restore_",
    "send_",
    "sort_",
    "start_",
    "trigger_",
    "update_",
    "upload_",
)
DESTRUCTIVE_PATTERNS = (
    "cancel_*",
    "delete_*",
    "remove_*",
    "wipe_*",
    "retire_*",
    "reboot_*",
    "shut_down_*",
    "dismiss_risky_user",
    "*application_password*",
    "*conditional_access_policy*",
    "*role_assignment*",
    "create_application",
    "create_service_principal",
    "update_organization",
    "update_admin_sharepoint",
)
ALWAYS_ALLOWED = {
    "health_check",
    "login",
    "login_power_platform",
    "login_windows_companion",
    "logout",
    "verify_login",
    "list_accounts",
    "search_tools",
    "searches",
    "run_hunting_query",
    "calendar_today",
}


def classify_tool_risk(tool_name: str) -> ToolRisk:
    """Classify a tool conservatively from its stable public name."""

    name = tool_name.lower().strip()
    if name in ALWAYS_ALLOWED:
        return ToolRisk.READ
    if any(fnmatch.fnmatch(name, pattern) for pattern in DESTRUCTIVE_PATTERNS):
        return ToolRisk.DESTRUCTIVE
    if name.startswith(READ_PREFIXES):
        return ToolRisk.READ
    if name.startswith(WRITE_PREFIXES):
        return ToolRisk.WRITE
    # Unknown actions fail into the write tier so generated tools cannot bypass
    # policy merely by choosing an unfamiliar verb.
    return ToolRisk.WRITE


class ToolPolicyDecision(BaseModel):
    """Auditable authorization decision for one MCP tool invocation."""

    model_config = ConfigDict(frozen=True)

    tool_name: str
    risk: ToolRisk
    allowed: bool
    reason: str


class MicrosoftToolPolicy:
    """Authorize tools using explicit deployment-level side-effect flags."""

    def __init__(self, settings: MicrosoftSettings | None = None):
        self.settings = settings or get_settings()

    def evaluate(self, tool_name: str) -> ToolPolicyDecision:
        """Return a deterministic, fail-closed decision for a tool."""

        risk = classify_tool_risk(tool_name)
        if risk is ToolRisk.READ:
            return ToolPolicyDecision(
                tool_name=tool_name,
                risk=risk,
                allowed=True,
                reason="read-only or authentication operation",
            )
        if risk is ToolRisk.WRITE and not self.settings.allow_write_tools:
            return ToolPolicyDecision(
                tool_name=tool_name,
                risk=risk,
                allowed=False,
                reason="MICROSOFT_ALLOW_WRITES is disabled",
            )
        if risk is ToolRisk.DESTRUCTIVE and not (
            self.settings.allow_write_tools and self.settings.allow_destructive_tools
        ):
            return ToolPolicyDecision(
                tool_name=tool_name,
                risk=risk,
                allowed=False,
                reason=(
                    "destructive operations require both MICROSOFT_ALLOW_WRITES "
                    "and MICROSOFT_ALLOW_DESTRUCTIVE"
                ),
            )
        return ToolPolicyDecision(
            tool_name=tool_name,
            risk=risk,
            allowed=True,
            reason="operation enabled by explicit deployment policy",
        )

    def require(self, tool_name: str) -> ToolPolicyDecision:
        """Raise a sanitized MCP error when the operation is denied."""

        decision = self.evaluate(tool_name)
        if not decision.allowed:
            raise ToolError(
                f"Tool '{tool_name}' is disabled by policy: {decision.reason}."
            )
        return decision


def _tool_name(message: Any) -> str | None:
    name = getattr(message, "name", None)
    if name:
        return str(name)
    params = getattr(message, "params", None)
    name = getattr(params, "name", None) if params is not None else None
    return str(name) if name else None


def _tool_action(message: Any) -> str | None:
    """Return the routed action from a FastMCP call without trusting its name."""

    params = getattr(message, "params", None)
    arguments = getattr(params, "arguments", None) if params is not None else None
    if not isinstance(arguments, Mapping):
        return None
    action = arguments.get("action")
    if not isinstance(action, str):
        return None
    normalized = action.strip()
    return normalized if normalized else None


class ToolPolicyMiddleware(Middleware):
    """Enforce side-effect policy before any MCP tool implementation runs."""

    def __init__(self, policy: MicrosoftToolPolicy | None = None):
        self.policy = policy or MicrosoftToolPolicy()

    async def on_call_tool(self, context: MiddlewareContext, call_next):
        name = _tool_name(context.message)
        if not name:
            raise ToolError("Unable to authorize unnamed Microsoft tool call.")
        action = _tool_action(context.message)
        # Condensed tools route many Graph operations through one stable MCP
        # name. Authorize the validated action verb, not merely the envelope,
        # so read operations remain usable while writes fail closed.
        policy_name = action if name.startswith("microsoft_") and action else name
        self.policy.require(policy_name)
        return await call_next(context)
