"""MCP tools for security operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_connector_sdk.mcp.action_dispatch import resolve_action
from agent_connector_sdk.mcp.concurrency import invoke_client_method
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client_dependency

_SECURITY_ACTIONS = (
    "list_security_alerts",
    "get_security_alert",
    "update_security_alert",
    "list_security_incidents",
    "get_security_incident",
    "update_security_incident",
    "list_secure_scores",
    "list_threat_intelligence_hosts",
    "get_threat_intelligence_host",
    "run_hunting_query",
    "list_risk_detections",
    "get_risk_detection",
    "list_risky_users",
    "get_risky_user",
    "dismiss_risky_user",
    "list_sensitivity_labels",
    "get_sensitivity_label",
)


def register_security_tools(mcp: FastMCP):
    @mcp.tool(tags={"security"})
    async def microsoft_security(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_security_alerts', 'get_security_alert', 'update_security_alert', 'list_security_incidents', 'get_security_incident', 'update_security_incident', 'list_secure_scores', 'list_threat_intelligence_hosts', 'get_threat_intelligence_host', 'run_hunting_query', 'list_risk_detections', 'get_risk_detection', 'list_risky_users', 'get_risky_user', 'dismiss_risky_user', 'list_sensitivity_labels', 'get_sensitivity_label'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client_dependency),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft security operations."""
        if ctx:
            await ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception:
            return {"error": "Invalid params_json"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        resolved = resolve_action(action, _SECURITY_ACTIONS, service="microsoft-agent")
        if isinstance(resolved, dict):
            return resolved
        action = resolved

        return await invoke_client_method(getattr(client, action), **kwargs)
