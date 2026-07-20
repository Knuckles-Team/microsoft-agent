#!/usr/bin/python
"""
Microsoft Agent A2A Supervisor Agent Server implementation.

CONCEPT:AU-ECO.mcp.fastmcp-middleware
"""

import logging
import sys
import warnings

from agent_utilities.agent.factory import create_agent_parser
from agent_utilities.core.config import setting
from agent_utilities.core.workspace import initialize_workspace
from agent_utilities.prompting.builder import (
    build_system_prompt_from_workspace,
    load_identity,
)
from agent_utilities.server import create_agent_server

from microsoft_agent._version import __version__

logger = logging.getLogger(__name__)


def agent_server():
    warnings.filterwarnings("ignore", message=".*urllib3.*or chardet.*")
    warnings.filterwarnings("ignore", category=DeprecationWarning, module="fastmcp")

    parser = create_agent_parser()
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler()],
    )

    initialize_workspace()
    meta = load_identity()
    agent_name = setting("DEFAULT_AGENT_NAME", meta.get("name", "Microsoft Agent"))
    system_prompt = setting(
        "AGENT_SYSTEM_PROMPT",
        meta.get("content") or build_system_prompt_from_workspace(),
    )
    print(f"{agent_name} v{__version__}", file=sys.stderr)

    create_agent_server(
        mcp_url=args.mcp_url,
        mcp_config=args.mcp_config,
        host=args.host,
        port=args.port,
        provider=args.provider,
        model_id=args.model_id,
        router_model=args.model_id,
        agent_model=args.model_id,
        base_url=args.base_url,
        api_key=args.api_key,
        custom_skills_directory=args.custom_skills_directory,
        name=agent_name,
        system_prompt=system_prompt,
        enable_web_ui=args.web,
        enable_otel=args.otel,
        otel_endpoint=args.otel_endpoint,
        otel_headers=args.otel_headers,
        otel_public_key=args.otel_public_key,
        otel_secret_key=args.otel_secret_key,
        otel_protocol=args.otel_protocol,
        debug=args.debug,
    )


if __name__ == "__main__":
    agent_server()
