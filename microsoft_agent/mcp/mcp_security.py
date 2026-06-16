"""MCP tools for security operations.

Auto-generated from mcp_server.py during ecosystem standardization.
"""

from agent_utilities.mcp_utilities import run_blocking
from fastmcp import Context, FastMCP
from fastmcp.dependencies import Depends
from pydantic import Field

from microsoft_agent.auth import get_client


def register_security_tools(mcp: FastMCP):
    @mcp.tool(tags={"security"})
    async def microsoft_security(
        action: str = Field(
            description="Action to perform. Must be one of: 'list_security_alerts', 'get_security_alert', 'update_security_alert', 'list_security_incidents', 'get_security_incident', 'update_security_incident', 'list_secure_scores', 'list_threat_intelligence_hosts', 'get_threat_intelligence_host', 'run_hunting_query', 'list_risk_detections', 'get_risk_detection', 'list_risky_users', 'get_risky_user', 'dismiss_risky_user', 'list_sensitivity_labels', 'get_sensitivity_label'"
        ),
        params_json: str = Field(
            default="{}", description="JSON string of parameters to pass to the action."
        ),
        client=Depends(get_client),
        ctx: Context | None = Field(
            default=None, description="MCP context for progress reporting"
        ),
    ) -> dict:
        """Manage microsoft security operations."""
        if ctx:
            ctx.info("Executing tool...")
        import json

        try:
            kwargs = json.loads(params_json)
        except Exception as e:
            return {"error": f"Invalid params_json: {e}"}

        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        if action == "list_security_alerts":
            return await run_blocking(client.list_security_alerts, **kwargs)
        if action == "get_security_alert":
            return await run_blocking(client.get_security_alert, **kwargs)
        if action == "update_security_alert":
            return await run_blocking(client.update_security_alert, **kwargs)
        if action == "list_security_incidents":
            return await run_blocking(client.list_security_incidents, **kwargs)
        if action == "get_security_incident":
            return await run_blocking(client.get_security_incident, **kwargs)
        if action == "update_security_incident":
            return await run_blocking(client.update_security_incident, **kwargs)
        if action == "list_secure_scores":
            return await run_blocking(client.list_secure_scores, **kwargs)
        if action == "list_threat_intelligence_hosts":
            return await run_blocking(client.list_threat_intelligence_hosts, **kwargs)
        if action == "get_threat_intelligence_host":
            return await run_blocking(client.get_threat_intelligence_host, **kwargs)
        if action == "run_hunting_query":
            return await run_blocking(client.run_hunting_query, **kwargs)
        if action == "list_risk_detections":
            return await run_blocking(client.list_risk_detections, **kwargs)
        if action == "get_risk_detection":
            return await run_blocking(client.get_risk_detection, **kwargs)
        if action == "list_risky_users":
            return await run_blocking(client.list_risky_users, **kwargs)
        if action == "get_risky_user":
            return await run_blocking(client.get_risky_user, **kwargs)
        if action == "dismiss_risky_user":
            return await run_blocking(client.dismiss_risky_user, **kwargs)
        if action == "list_sensitivity_labels":
            return await run_blocking(client.list_sensitivity_labels, **kwargs)
        if action == "get_sensitivity_label":
            return await run_blocking(client.get_sensitivity_label, **kwargs)
        raise ValueError(f"Unknown action: {action}")
