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
from pathlib import Path
from typing import Any

from scion_sdk.config import load_config
from scion_sdk.mcp.concurrency import invoke_client_method
from scion_sdk.mcp.content import ConnectorContent
from scion_sdk.mcp.server import create_mcp_server
from scion_sdk.mcp.tool_surface import register_tool_surface
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


def register_extra_surface(mcp: FastMCP) -> None:
    """Register the tools beyond the Graph operations, and the side-effect policy.

    Shared by the stand-alone server and the declared connector so both serve
    the same tools under the same fail-closed authorization.
    """
    register_kg_tools(mcp)

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
    mcp.add_middleware(ToolPolicyMiddleware(MicrosoftToolPolicy(settings)))


def get_mcp_instance() -> tuple[Any, ...]:
    """Initialize and return the MCP instance."""
    load_config()
    args, mcp, middlewares = create_mcp_server(
        name="microsoft-agent MCP",
        version=__version__,
        instructions="microsoft-agent MCP Server — one typed tool per Microsoft Graph operation.",
        content=ConnectorContent.for_package(
            "microsoft_agent",
            Path(__file__).resolve().parent,
            __version__,
        ),
    )

    @mcp.custom_route("/health", methods=["GET"])
    async def health_check(request: Request) -> JSONResponse:
        return JSONResponse({"status": "OK"})

    register_tool_surface(
        mcp,
        client_cls=MicrosoftGraphApi,
        get_client=get_client_dependency,
        service="microsoft-agent",
    )
    register_extra_surface(mcp)

    for mw in middlewares:
        mcp.add_middleware(mw)
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
        elif args.transport == "sse":
            mcp.run(transport="sse", host=args.host, port=args.port)
        else:
            logger.error("Invalid transport", extra={"transport": args.transport})
            sys.exit(1)
    finally:
        clear_integration_client_caches()


if __name__ == "__main__":
    mcp_server()
