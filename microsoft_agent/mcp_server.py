#!/usr/bin/python
"""
Microsoft Graph MCP Server implementation.
"""

import warnings

from fastmcp import FastMCP
from fastmcp.dependencies import Depends
from fastmcp.utilities.logging import get_logger
from pydantic import Field

# Filter RequestsDependencyWarning early to prevent log spam
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    try:
        from requests.exceptions import RequestsDependencyWarning

        warnings.filterwarnings("ignore", category=RequestsDependencyWarning)
    except ImportError:
        pass

warnings.filterwarnings("ignore", message=".*urllib3.*or chardet.*")
warnings.filterwarnings("ignore", message=".*urllib3.*or charset_normalizer.*")

import logging
import sys
from typing import Any

from agent_utilities.core.config import load_config
from agent_utilities.mcp.concurrency import invoke_client_method
from agent_utilities.mcp.server_factory import create_mcp_server
from agent_utilities.mcp.verbose_tools import register_tool_surface
from starlette.requests import Request
from starlette.responses import JSONResponse

from microsoft_agent._version import __version__
from microsoft_agent.api_client import MicrosoftGraphApi
from microsoft_agent.auth import get_client_dependency
from microsoft_agent.integration_tools import (
    clear_integration_client_caches,
    register_document_tools,
    register_intune_tools,
    register_power_platform_tools,
    register_windows_companion_tools,
)
from microsoft_agent.mcp import (
    register_admin_tools as register_admin_tools,
)
from microsoft_agent.mcp import (
    register_agreements_tools as register_agreements_tools,
)
from microsoft_agent.mcp import (
    register_applications_tools as register_applications_tools,
)
from microsoft_agent.mcp import (
    register_audit_tools as register_audit_tools,
)
from microsoft_agent.mcp import (
    register_auth_tools as register_auth_tools,
)
from microsoft_agent.mcp import (
    register_calendar_tools as register_calendar_tools,
)
from microsoft_agent.mcp import (
    register_chat_tools as register_chat_tools,
)
from microsoft_agent.mcp import (
    register_communications_tools as register_communications_tools,
)
from microsoft_agent.mcp import (
    register_connections_tools as register_connections_tools,
)
from microsoft_agent.mcp import (
    register_contacts_tools as register_contacts_tools,
)
from microsoft_agent.mcp import (
    register_devices_tools as register_devices_tools,
)
from microsoft_agent.mcp import (
    register_directory_tools as register_directory_tools,
)
from microsoft_agent.mcp import (
    register_domains_tools as register_domains_tools,
)
from microsoft_agent.mcp import (
    register_education_tools as register_education_tools,
)
from microsoft_agent.mcp import (
    register_employee_experience_tools as register_employee_experience_tools,
)
from microsoft_agent.mcp import (
    register_files_tools as register_files_tools,
)
from microsoft_agent.mcp import (
    register_groups_tools as register_groups_tools,
)
from microsoft_agent.mcp import (
    register_identity_tools as register_identity_tools,
)
from microsoft_agent.mcp import (
    register_mail_tools as register_mail_tools,
)
from microsoft_agent.mcp import (
    register_meta_tools as register_meta_tools,
)
from microsoft_agent.mcp import (
    register_notes_tools as register_notes_tools,
)
from microsoft_agent.mcp import (
    register_organization_tools as register_organization_tools,
)
from microsoft_agent.mcp import (
    register_places_tools as register_places_tools,
)
from microsoft_agent.mcp import (
    register_policies_tools as register_policies_tools,
)
from microsoft_agent.mcp import (
    register_print_tools as register_print_tools,
)
from microsoft_agent.mcp import (
    register_privacy_tools as register_privacy_tools,
)
from microsoft_agent.mcp import (
    register_reports_tools as register_reports_tools,
)
from microsoft_agent.mcp import (
    register_search_tools as register_search_tools,
)
from microsoft_agent.mcp import (
    register_security_tools as register_security_tools,
)
from microsoft_agent.mcp import (
    register_sites_tools as register_sites_tools,
)
from microsoft_agent.mcp import (
    register_solutions_tools as register_solutions_tools,
)
from microsoft_agent.mcp import (
    register_storage_tools as register_storage_tools,
)
from microsoft_agent.mcp import (
    register_subscriptions_tools as register_subscriptions_tools,
)
from microsoft_agent.mcp import (
    register_tasks_tools as register_tasks_tools,
)
from microsoft_agent.mcp import (
    register_teams_tools as register_teams_tools,
)
from microsoft_agent.mcp import (
    register_user_tools as register_user_tools,
)
from microsoft_agent.office_bridge import register_office_bridge
from microsoft_agent.settings import get_settings
from microsoft_agent.tool_policy import MicrosoftToolPolicy, ToolPolicyMiddleware

logger = get_logger(name="microsoft-agent")
logger.setLevel(logging.INFO)


def register_kg_tools(mcp: FastMCP):
    import json as _json

    from microsoft_agent import kg_ingest

    _projection_select = {
        "messages": "id,from,toRecipients",
        "events": "id,organizer,attendees",
        "files": "id,createdBy",
        "users": "id",
    }

    def _bounded_params(kind: str, params_json: str) -> dict[str, Any]:
        if kind not in _projection_select:
            raise ValueError("Unknown Microsoft projection kind")
        try:
            params = _json.loads(params_json) if params_json else {}
        except (TypeError, ValueError) as exc:
            raise ValueError("Invalid params_json") from exc
        if not isinstance(params, dict):
            raise ValueError("params_json must contain an object")
        unsupported = set(params) - {"$top", "$filter", "$orderby", "$skiptoken"}
        if unsupported:
            raise ValueError("params_json contains unsupported projection options")
        try:
            top = int(params.get("$top", 100))
        except (TypeError, ValueError) as exc:
            raise ValueError("$top must be an integer") from exc
        if not 1 <= top <= 100:
            raise ValueError("$top must be between 1 and 100")
        return {**params, "$top": top, "$select": _projection_select[kind]}

    async def _load_records(
        kind: str,
        params_json: str,
        client: MicrosoftGraphApi,
        *,
        drive_id: str | None = None,
        drive_item_id: str | None = None,
    ) -> list[dict[str, Any]]:
        listers = {
            "messages": client.list_mail_messages,
            "events": client.list_calendar_events,
            "users": client.list_users,
        }
        params = _bounded_params(kind, params_json)
        if kind == "files":
            if not drive_id or not drive_item_id:
                raise ValueError("File projection requires drive_id and drive_item_id")
            response = await invoke_client_method(
                client.list_folder_files,
                drive_id,
                drive_item_id,
                params=params,
            )
            return kg_ingest._records(response)
        lister = listers.get(kind)
        if lister is None:
            raise ValueError("Unknown Microsoft projection kind")
        response = await invoke_client_method(lister, params=params)
        return kg_ingest._records(response)

    @mcp.tool(tags={"kg", "read"})
    async def list_microsoft_ingestion_projection(
        kind: str = Field(
            default="messages",
            description="Projection kind: messages, events, files, or users.",
        ),
        params_json: str = Field(
            default="{}",
            description="Bounded JSON object containing approved OData options.",
        ),
        drive_id: str | None = Field(
            default=None,
            description="Externally configured drive identifier for file projection.",
        ),
        drive_item_id: str | None = Field(
            default=None,
            description="Externally configured folder identifier for file projection.",
        ),
        client=Depends(get_client_dependency),
    ) -> dict[str, Any]:
        """Return only keyed opaque nodes and structural relationships.

        Raw Microsoft identifiers and content never appear in the returned source
        connector projection. The deployment-owned pseudonymization key is required.
        """

        records = await _load_records(
            kind,
            params_json,
            client,
            drive_id=drive_id,
            drive_item_id=drive_item_id,
        )
        projection = kg_ingest.project_records(kind, records)
        return {"kind": kind, "listed": len(records), **projection}

    return None


def get_mcp_instance() -> tuple[Any, ...]:
    """Initialize and return the MCP instance."""
    load_config()
    args, mcp, middlewares = create_mcp_server(
        name="microsoft-agent MCP",
        version=__version__,
        instructions="microsoft-agent MCP Server — Condensed Action-Routed Tools.",
    )

    @mcp.custom_route("/health", methods=["GET"])
    async def health_check(request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    register_tool_surface(
        mcp,
        client_cls=MicrosoftGraphApi,
        get_client=get_client_dependency,
        service="microsoft-agent",
        tools_module=sys.modules[__name__],
    )

    settings = get_settings()
    if settings.tool_group_enabled("documents"):
        register_document_tools(mcp)
        register_office_bridge(mcp)
    if settings.tool_group_enabled("power_platform"):
        register_power_platform_tools(mcp)
    if settings.tool_group_enabled("windows"):
        register_windows_companion_tools(mcp)
    if settings.tool_group_enabled("intune"):
        register_intune_tools(mcp)

    for mw in middlewares:
        mcp.add_middleware(mw)
    mcp.add_middleware(ToolPolicyMiddleware(MicrosoftToolPolicy(settings)))
    return mcp, args, middlewares


def mcp_server() -> None:
    mcp, args, middlewares = get_mcp_instance()
    print(f"microsoft-agent MCP v{__version__}", file=sys.stderr)
    print("\nStarting MCP Server", file=sys.stderr)
    print(f"  Transport: {args.transport.upper()}", file=sys.stderr)
    print(f"  Auth: {args.auth_type}", file=sys.stderr)

    try:
        if args.transport == "stdio":
            mcp.run(transport="stdio")
        elif args.transport == "streamable-http":
            mcp.run(transport="streamable-http", host=args.host, port=args.port)
        else:
            logger.error("Invalid transport", extra={"transport": args.transport})
            sys.exit(1)
    finally:
        clear_integration_client_caches()


if __name__ == "__main__":
    mcp_server()
