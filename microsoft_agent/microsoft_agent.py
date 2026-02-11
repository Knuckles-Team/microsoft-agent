#!/usr/bin/python
import sys

# coding: utf-8
import json
import os
import argparse
import logging
import inspect
import uvicorn
from typing import Optional, Any, List, Dict
from contextlib import asynccontextmanager

from pydantic_ai import Agent, ModelSettings, RunContext
from pydantic_ai.mcp import (
    load_mcp_servers,
    MCPServerStreamableHTTP,
    MCPServerSSE,
)
from pydantic_ai_skills import SkillsToolset
from fasta2a import Skill
from microsoft_agent.utils import (
    to_integer,
    to_boolean,
    to_float,
    to_list,
    to_dict,
    get_mcp_config_path,
    get_skills_path,
    load_skills_from_directory,
    create_model,
    tool_in_tag,
    prune_large_messages,
)
import microsoft_agent.microsoft_mcp as microsoft_mcp
from fastapi import FastAPI, Request
from starlette.responses import Response, StreamingResponse
from pydantic import ValidationError
from pydantic_ai.ui import SSE_CONTENT_TYPE
from pydantic_ai.ui.ag_ui import AGUIAdapter

__version__ = "0.2.4"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],  # Output to console
)
logging.getLogger("pydantic_ai").setLevel(logging.INFO)
logging.getLogger("fastmcp").setLevel(logging.INFO)
logging.getLogger("httpx").setLevel(logging.INFO)
logger = logging.getLogger(__name__)

DEFAULT_HOST = os.getenv("HOST", "0.0.0.0")
DEFAULT_PORT = to_integer(string=os.getenv("PORT", "9000"))
DEFAULT_DEBUG = to_boolean(string=os.getenv("DEBUG", "False"))
DEFAULT_PROVIDER = os.getenv("PROVIDER", "openai")
DEFAULT_MODEL_ID = os.getenv("MODEL_ID", "qwen/qwen3-coder-next")
DEFAULT_LLM_BASE_URL = os.getenv("LLM_BASE_URL", "http://host.docker.internal:1234/v1")
DEFAULT_LLM_API_KEY = os.getenv("LLM_API_KEY", "ollama")
DEFAULT_MCP_URL = os.getenv("MCP_URL", None)
DEFAULT_MCP_CONFIG = os.getenv("MCP_CONFIG", get_mcp_config_path())
DEFAULT_SKILLS_DIRECTORY = os.getenv("SKILLS_DIRECTORY", get_skills_path())
DEFAULT_ENABLE_WEB_UI = to_boolean(os.getenv("ENABLE_WEB_UI", "False"))
DEFAULT_SSL_VERIFY = to_boolean(os.getenv("SSL_VERIFY", "True"))

# Model Settings
DEFAULT_MAX_TOKENS = to_integer(os.getenv("MAX_TOKENS", "16384"))
DEFAULT_TEMPERATURE = to_float(os.getenv("TEMPERATURE", "0.7"))
DEFAULT_TOP_P = to_float(os.getenv("TOP_P", "1.0"))
DEFAULT_TIMEOUT = to_float(os.getenv("TIMEOUT", "32400.0"))
DEFAULT_TOOL_TIMEOUT = to_float(os.getenv("TOOL_TIMEOUT", "32400.0"))
DEFAULT_PARALLEL_TOOL_CALLS = to_boolean(os.getenv("PARALLEL_TOOL_CALLS", "True"))
DEFAULT_SEED = to_integer(os.getenv("SEED", None))
DEFAULT_PRESENCE_PENALTY = to_float(os.getenv("PRESENCE_PENALTY", "0.0"))
DEFAULT_FREQUENCY_PENALTY = to_float(os.getenv("FREQUENCY_PENALTY", "0.0"))
DEFAULT_LOGIT_BIAS = to_dict(os.getenv("LOGIT_BIAS", None))
DEFAULT_STOP_SEQUENCES = to_list(os.getenv("STOP_SEQUENCES", None))
DEFAULT_EXTRA_HEADERS = to_dict(os.getenv("EXTRA_HEADERS", None))
DEFAULT_EXTRA_BODY = to_dict(os.getenv("EXTRA_BODY", None))

AGENT_NAME = "MicrosoftAgent"
AGENT_DESCRIPTION = "A multi-agent system for managing Microsoft Graph resources via delegated specialists."
SUPERVISOR_SYSTEM_PROMPT = os.environ.get(
    "SUPERVISOR_SYSTEM_PROMPT",
    default=(
        "You are the Microsoft Graph Supervisor Agent.\n"
        "Your goal is to assist the user by assigning tasks to specialized child agents through your available toolset.\n"
        "Analyze the user's request and determine which domain(s) it falls into (e.g., mail, calendar, drive, users, etc.).\n"
        "Then, call the appropriate tool(s) to delegate the task.\n"
        "Synthesize the results from the child agents into a final helpful response.\n"
        "Always be warm, professional, and helpful."
        "Note: The final response should contain all the relevant information from the tool executions. Never leave out any relevant information or leave it to the user to find it. "
        "You are the final authority on the user's request and the final communicator to the user. Present information as logically and concisely as possible. "
        "Explore using organized output with headers, sections, lists, and tables to make the information easy to navigate. "
        "If there are gaps in the information, clearly state that information is missing. Do not make assumptions or invent placeholder information, only use the information which is available. "
        "Plainly say that you do not have that information. If you were given an error output, try to capture as many relevant details as possible from the error output and include it in your response as a bug formatted pull request."
    ),
)

# -------------------------------------------------------------------------
# 1. System Prompts
# -------------------------------------------------------------------------

ACTIVITIESCONTAINER_AGENT_PROMPT = "You are the Activitiescontainer Agent. You manage activitiescontainer resources using the available tools."
APPLICATION_AGENT_PROMPT = "You are the Application Agent. You manage application resources using the available tools."
CALENDAR_AGENT_PROMPT = "You are the Calendar Agent. You manage calendar resources using the available tools."
CHART_AGENT_PROMPT = (
    "You are the Chart Agent. You manage chart resources using the available tools."
)
CHARTCOLLECTION_AGENT_PROMPT = "You are the Chartcollection Agent. You manage chartcollection resources using the available tools."
CONVERSATIONTHREAD_AGENT_PROMPT = "You are the Conversationthread Agent. You manage conversationthread resources using the available tools."
DIRECTORY_AGENT_PROMPT = "You are the Directory Agent. You manage directory resources using the available tools."
DIRECTORYOBJECT_AGENT_PROMPT = "You are the Directoryobject Agent. You manage directoryobject resources using the available tools."
DRIVE_AGENT_PROMPT = (
    "You are the Drive Agent. You manage drive resources using the available tools."
)
DRIVEITEM_AGENT_PROMPT = "You are the Driveitem Agent. You manage driveitem resources using the available tools."
EVENT_AGENT_PROMPT = (
    "You are the Event Agent. You manage event resources using the available tools."
)
GROUP_AGENT_PROMPT = (
    "You are the Group Agent. You manage group resources using the available tools."
)
ITEMACTIVITYSTAT_AGENT_PROMPT = "You are the Itemactivitystat Agent. You manage itemactivitystat resources using the available tools."
ITEMANALYTICS_AGENT_PROMPT = "You are the Itemanalytics Agent. You manage itemanalytics resources using the available tools."
NOTEBOOK_AGENT_PROMPT = "You are the Notebook Agent. You manage notebook resources using the available tools."
ONENOTESECTION_AGENT_PROMPT = "You are the Onenotesection Agent. You manage onenotesection resources using the available tools."
OPENTYPEEXTENSION_AGENT_PROMPT = "You are the Opentypeextension Agent. You manage opentypeextension resources using the available tools."
PERMISSION_AGENT_PROMPT = "You are the Permission Agent. You manage permission resources using the available tools."
PLANNERGROUP_AGENT_PROMPT = "You are the Plannergroup Agent. You manage plannergroup resources using the available tools."
PLANNERUSER_AGENT_PROMPT = "You are the Planneruser Agent. You manage planneruser resources using the available tools."
REPORTROOT_AGENT_PROMPT = "You are the Reportroot Agent. You manage reportroot resources using the available tools."
RESOURCES_BROWSER_AGENT_PROMPT = "You are the Resources_browser Agent. You manage resources_browser resources using the available tools."
RESOURCES_CALENDAR_AGENT_PROMPT = "You are the Resources_calendar Agent. You manage resources_calendar resources using the available tools."
RESOURCES_CALLRECORDS_AGENT_PROMPT = "You are the Resources_callrecords Agent. You manage resources_callrecords resources using the available tools."
RESOURCES_CHANGE_AGENT_PROMPT = "You are the Resources_change Agent. You manage resources_change resources using the available tools."
RESOURCES_GROUPS_AGENT_PROMPT = "You are the Resources_groups Agent. You manage resources_groups resources using the available tools."
RESOURCES_ONENOTE_AGENT_PROMPT = "You are the Resources_onenote Agent. You manage resources_onenote resources using the available tools."
RESOURCES_PLANNER_AGENT_PROMPT = "You are the Resources_planner Agent. You manage resources_planner resources using the available tools."
SECTION_AGENT_PROMPT = (
    "You are the Section Agent. You manage section resources using the available tools."
)
SERVICEPRINCIPAL_AGENT_PROMPT = "You are the Serviceprincipal Agent. You manage serviceprincipal resources using the available tools."
SUBSCRIPTION_AGENT_PROMPT = "You are the Subscription Agent. You manage subscription resources using the available tools."
SUBSCRIPTIONS_AGENT_PROMPT = "You are the Subscriptions Agent. You manage subscriptions resources using the available tools."
USER_AGENT_PROMPT = (
    "You are the User Agent. You manage user resources using the available tools."
)
USERDATASECURITYANDGOVERNANCE_AGENT_PROMPT = "You are the Userdatasecurityandgovernance Agent. You manage userdatasecurityandgovernance resources using the available tools."
USERPROTECTIONSCOPECONTAINER_AGENT_PROMPT = "You are the Userprotectionscopecontainer Agent. You manage userprotectionscopecontainer resources using the available tools."

# -------------------------------------------------------------------------
# 2. Agent Creation Logic
# -------------------------------------------------------------------------


def create_agent(
    provider: str = DEFAULT_PROVIDER,
    model_id: str = DEFAULT_MODEL_ID,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    mcp_url: str = DEFAULT_MCP_URL,
    mcp_config: str = DEFAULT_MCP_CONFIG,
    skills_directory: Optional[str] = DEFAULT_SKILLS_DIRECTORY,
    ssl_verify: bool = DEFAULT_SSL_VERIFY,
) -> Agent:
    """
    Creates the Supervisor Agent with sub-agents registered as tools.
    """
    logger.info("Initializing Multi-Agent System for Microsoft...")

    model = create_model(
        provider=provider,
        model_id=model_id,
        base_url=base_url,
        api_key=api_key,
        ssl_verify=ssl_verify,
    )
    settings = ModelSettings(
        max_tokens=DEFAULT_MAX_TOKENS,
        temperature=DEFAULT_TEMPERATURE,
        top_p=DEFAULT_TOP_P,
        timeout=DEFAULT_TIMEOUT,
        parallel_tool_calls=DEFAULT_PARALLEL_TOOL_CALLS,
        seed=DEFAULT_SEED,
        presence_penalty=DEFAULT_PRESENCE_PENALTY,
        frequency_penalty=DEFAULT_FREQUENCY_PENALTY,
        logit_bias=DEFAULT_LOGIT_BIAS,
        stop_sequences=DEFAULT_STOP_SEQUENCES,
        extra_headers=DEFAULT_EXTRA_HEADERS,
        extra_body=DEFAULT_EXTRA_BODY,
    )

    # Load master toolsets
    agent_toolsets = []
    if mcp_url:
        if "sse" in mcp_url.lower():
            server = MCPServerSSE(mcp_url)
        else:
            server = MCPServerStreamableHTTP(mcp_url)
        agent_toolsets.append(server)
        logger.info(f"Connected to MCP Server: {mcp_url}")
    elif mcp_config:
        mcp_toolset = load_mcp_servers(mcp_config)
        agent_toolsets.extend(mcp_toolset)
        logger.info(f"Connected to MCP Config JSON: {mcp_toolset}")

    if skills_directory and os.path.exists(skills_directory):
        agent_toolsets.append(SkillsToolset(directories=[str(skills_directory)]))

    # Identify local tools from microsoft_mcp
    resource_tools: Dict[str, List] = {}
    for name, func in inspect.getmembers(microsoft_mcp):
        if hasattr(func, "__call__") and getattr(func, "__module__", "").endswith(
            "microsoft_mcp"
        ):
            # Assuming tool names are like action_resource or just resource_action
            # The previous code used split('_')[-1] as the resource name
            parts = name.split("_")
            if len(parts) > 1:
                res = parts[-1]
                if res not in resource_tools:
                    resource_tools[res] = []
                resource_tools[res].append(func)

    # Define Tag -> Prompt map
    agent_defs = {
        "activitiescontainer": (
            ACTIVITIESCONTAINER_AGENT_PROMPT,
            "Microsoft_Activitiescontainer_Agent",
        ),
        "application": (APPLICATION_AGENT_PROMPT, "Microsoft_Application_Agent"),
        "calendar": (CALENDAR_AGENT_PROMPT, "Microsoft_Calendar_Agent"),
        "chart": (CHART_AGENT_PROMPT, "Microsoft_Chart_Agent"),
        "chartcollection": (
            CHARTCOLLECTION_AGENT_PROMPT,
            "Microsoft_Chartcollection_Agent",
        ),
        "conversationthread": (
            CONVERSATIONTHREAD_AGENT_PROMPT,
            "Microsoft_Conversationthread_Agent",
        ),
        "directory": (DIRECTORY_AGENT_PROMPT, "Microsoft_Directory_Agent"),
        "directoryobject": (
            DIRECTORYOBJECT_AGENT_PROMPT,
            "Microsoft_Directoryobject_Agent",
        ),
        "drive": (DRIVE_AGENT_PROMPT, "Microsoft_Drive_Agent"),
        "driveitem": (DRIVEITEM_AGENT_PROMPT, "Microsoft_Driveitem_Agent"),
        "event": (EVENT_AGENT_PROMPT, "Microsoft_Event_Agent"),
        "group": (GROUP_AGENT_PROMPT, "Microsoft_Group_Agent"),
        "itemactivitystat": (
            ITEMACTIVITYSTAT_AGENT_PROMPT,
            "Microsoft_Itemactivitystat_Agent",
        ),
        "itemanalytics": (ITEMANALYTICS_AGENT_PROMPT, "Microsoft_Itemanalytics_Agent"),
        "notebook": (NOTEBOOK_AGENT_PROMPT, "Microsoft_Notebook_Agent"),
        "onenotesection": (
            ONENOTESECTION_AGENT_PROMPT,
            "Microsoft_Onenotesection_Agent",
        ),
        "opentypeextension": (
            OPENTYPEEXTENSION_AGENT_PROMPT,
            "Microsoft_Opentypeextension_Agent",
        ),
        "permission": (PERMISSION_AGENT_PROMPT, "Microsoft_Permission_Agent"),
        "plannergroup": (PLANNERGROUP_AGENT_PROMPT, "Microsoft_Plannergroup_Agent"),
        "planneruser": (PLANNERUSER_AGENT_PROMPT, "Microsoft_Planneruser_Agent"),
        "reportroot": (REPORTROOT_AGENT_PROMPT, "Microsoft_Reportroot_Agent"),
        "resources_browser": (
            RESOURCES_BROWSER_AGENT_PROMPT,
            "Microsoft_Resources_Browser_Agent",
        ),
        "resources_calendar": (
            RESOURCES_CALENDAR_AGENT_PROMPT,
            "Microsoft_Resources_Calendar_Agent",
        ),
        "resources_callrecords": (
            RESOURCES_CALLRECORDS_AGENT_PROMPT,
            "Microsoft_Resources_Callrecords_Agent",
        ),
        "resources_change": (
            RESOURCES_CHANGE_AGENT_PROMPT,
            "Microsoft_Resources_Change_Agent",
        ),
        "resources_groups": (
            RESOURCES_GROUPS_AGENT_PROMPT,
            "Microsoft_Resources_Groups_Agent",
        ),
        "resources_onenote": (
            RESOURCES_ONENOTE_AGENT_PROMPT,
            "Microsoft_Resources_Onenote_Agent",
        ),
        "resources_planner": (
            RESOURCES_PLANNER_AGENT_PROMPT,
            "Microsoft_Resources_Planner_Agent",
        ),
        "section": (SECTION_AGENT_PROMPT, "Microsoft_Section_Agent"),
        "serviceprincipal": (
            SERVICEPRINCIPAL_AGENT_PROMPT,
            "Microsoft_Serviceprincipal_Agent",
        ),
        "subscription": (SUBSCRIPTION_AGENT_PROMPT, "Microsoft_Subscription_Agent"),
        "subscriptions": (SUBSCRIPTIONS_AGENT_PROMPT, "Microsoft_Subscriptions_Agent"),
        "user": (USER_AGENT_PROMPT, "Microsoft_User_Agent"),
        "userdatasecurityandgovernance": (
            USERDATASECURITYANDGOVERNANCE_AGENT_PROMPT,
            "Microsoft_Userdatasecurityandgovernance_Agent",
        ),
        "userprotectionscopecontainer": (
            USERPROTECTIONSCOPECONTAINER_AGENT_PROMPT,
            "Microsoft_Userprotectionscopecontainer_Agent",
        ),
    }

    child_agents = {}

    for tag, (system_prompt, agent_name) in agent_defs.items():
        tag_toolsets = []
        # Filter external MCP tools by tag
        for ts in agent_toolsets:

            def filter_func(ctx, tool_def, t=tag):
                return tool_in_tag(tool_def, t)

            if hasattr(ts, "filtered"):
                filtered_ts = ts.filtered(filter_func)
                tag_toolsets.append(filtered_ts)
            else:
                pass

        # Add local tools
        local_tools = resource_tools.get(tag, [])

        agent = Agent(
            model=model,
            system_prompt=system_prompt,
            name=agent_name,
            toolsets=tag_toolsets,
            tool_timeout=DEFAULT_TOOL_TIMEOUT,
            model_settings=settings,
        )

        # Register local tools directly
        for t in local_tools:
            agent.tool(t)

        child_agents[tag] = agent

    # Create Supervisor
    supervisor = Agent(
        model=model,
        system_prompt=SUPERVISOR_SYSTEM_PROMPT,
        model_settings=settings,
        name=AGENT_NAME,
        deps_type=Any,
    )

    # Define delegation tools

    @supervisor.tool
    async def assign_task_to_activitiescontainer_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to activitiescontainer to the Activitiescontainer Agent."""
        return (
            await child_agents["activitiescontainer"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_application_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to application to the Application Agent."""
        return (
            await child_agents["application"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_calendar_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to calendar to the Calendar Agent."""
        return (
            await child_agents["calendar"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_chart_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to chart to the Chart Agent."""
        return (
            await child_agents["chart"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_chartcollection_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to chartcollection to the Chartcollection Agent."""
        return (
            await child_agents["chartcollection"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_conversationthread_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to conversationthread to the Conversationthread Agent."""
        return (
            await child_agents["conversationthread"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_directory_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to directory to the Directory Agent."""
        return (
            await child_agents["directory"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_directoryobject_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to directoryobject to the Directoryobject Agent."""
        return (
            await child_agents["directoryobject"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_drive_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to drive to the Drive Agent."""
        return (
            await child_agents["drive"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_driveitem_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to driveitem to the Driveitem Agent."""
        return (
            await child_agents["driveitem"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_event_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to event to the Event Agent."""
        return (
            await child_agents["event"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_group_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to group to the Group Agent."""
        return (
            await child_agents["group"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_itemactivitystat_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to itemactivitystat to the Itemactivitystat Agent."""
        return (
            await child_agents["itemactivitystat"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_itemanalytics_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to itemanalytics to the Itemanalytics Agent."""
        return (
            await child_agents["itemanalytics"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_notebook_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to notebook to the Notebook Agent."""
        return (
            await child_agents["notebook"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_onenotesection_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to onenotesection to the Onenotesection Agent."""
        return (
            await child_agents["onenotesection"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_opentypeextension_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to opentypeextension to the Opentypeextension Agent."""
        return (
            await child_agents["opentypeextension"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_permission_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to permission to the Permission Agent."""
        return (
            await child_agents["permission"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_plannergroup_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to plannergroup to the Plannergroup Agent."""
        return (
            await child_agents["plannergroup"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_planneruser_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to planneruser to the Planneruser Agent."""
        return (
            await child_agents["planneruser"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_reportroot_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to reportroot to the Reportroot Agent."""
        return (
            await child_agents["reportroot"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_resources_browser_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_browser to the Resources_browser Agent."""
        return (
            await child_agents["resources_browser"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_calendar_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_calendar to the Resources_calendar Agent."""
        return (
            await child_agents["resources_calendar"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_callrecords_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_callrecords to the Resources_callrecords Agent."""
        return (
            await child_agents["resources_callrecords"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_change_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_change to the Resources_change Agent."""
        return (
            await child_agents["resources_change"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_groups_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_groups to the Resources_groups Agent."""
        return (
            await child_agents["resources_groups"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_onenote_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_onenote to the Resources_onenote Agent."""
        return (
            await child_agents["resources_onenote"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_resources_planner_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to resources_planner to the Resources_planner Agent."""
        return (
            await child_agents["resources_planner"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_section_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to section to the Section Agent."""
        return (
            await child_agents["section"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_serviceprincipal_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to serviceprincipal to the Serviceprincipal Agent."""
        return (
            await child_agents["serviceprincipal"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_subscription_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to subscription to the Subscription Agent."""
        return (
            await child_agents["subscription"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_subscriptions_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to subscriptions to the Subscriptions Agent."""
        return (
            await child_agents["subscriptions"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_user_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to user to the User Agent."""
        return (
            await child_agents["user"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_userdatasecurityandgovernance_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to userdatasecurityandgovernance to the Userdatasecurityandgovernance Agent."""
        return (
            await child_agents["userdatasecurityandgovernance"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_userprotectionscopecontainer_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to userprotectionscopecontainer to the Userprotectionscopecontainer Agent."""
        return (
            await child_agents["userprotectionscopecontainer"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    return supervisor


async def chat(agent: Agent, prompt: str):
    result = await agent.run(prompt)
    print(f"Response:\n\n{result.output}")


async def node_chat(agent: Agent, prompt: str) -> List:
    nodes = []
    async with agent.iter(prompt) as agent_run:
        async for node in agent_run:
            nodes.append(node)
            print(node)
    return nodes


async def stream_chat(agent: Agent, prompt: str) -> None:
    async with agent.run_stream(prompt) as result:
        async for text_chunk in result.stream_text(delta=True):
            print(text_chunk, end="", flush=True)
        print("\nDone!")


def create_agent_server(
    provider: str = DEFAULT_PROVIDER,
    model_id: str = DEFAULT_MODEL_ID,
    base_url: Optional[str] = None,
    api_key: Optional[str] = None,
    mcp_url: str = DEFAULT_MCP_URL,
    mcp_config: str = DEFAULT_MCP_CONFIG,
    skills_directory: Optional[str] = DEFAULT_SKILLS_DIRECTORY,
    debug: Optional[bool] = DEFAULT_DEBUG,
    host: Optional[str] = DEFAULT_HOST,
    port: Optional[int] = DEFAULT_PORT,
    enable_web_ui: bool = DEFAULT_ENABLE_WEB_UI,
    ssl_verify: bool = DEFAULT_SSL_VERIFY,
):
    print(
        f"Starting {AGENT_NAME} with provider={provider}, model={model_id}, mcp={mcp_url} | {mcp_config}"
    )
    agent = create_agent(
        provider=provider,
        model_id=model_id,
        base_url=base_url,
        api_key=api_key,
        mcp_url=mcp_url,
        mcp_config=mcp_config,
        skills_directory=skills_directory,
        ssl_verify=ssl_verify,
    )

    if skills_directory and os.path.exists(skills_directory):
        skills = load_skills_from_directory(skills_directory)
        logger.info(f"Loaded {len(skills)} skills from {skills_directory}")
    else:
        skills = [
            Skill(
                id="microsoft_agent",
                name="Microsoft Agent",
                description="This Microsoft skill grants access to all Microsoft Graph tools provided by the Microsoft MCP Server",
                tags=["microsoft"],
                input_modes=["text"],
                output_modes=["text"],
            )
        ]

    a2a_app = agent.to_a2a(
        name=AGENT_NAME,
        description=AGENT_DESCRIPTION,
        version=__version__,
        skills=skills,
        debug=debug,
    )

    @asynccontextmanager
    async def lifespan(app: FastAPI):
        if hasattr(a2a_app, "router") and hasattr(a2a_app.router, "lifespan_context"):
            async with a2a_app.router.lifespan_context(a2a_app):
                yield
        else:
            yield

    app = FastAPI(
        title=f"{AGENT_NAME} - A2A + AG-UI Server",
        description=AGENT_DESCRIPTION,
        debug=debug,
        lifespan=lifespan,
    )

    @app.get("/health")
    async def health_check():
        return {"status": "OK"}

    app.mount("/a2a", a2a_app)

    @app.post("/ag-ui")
    async def ag_ui_endpoint(request: Request) -> Response:
        accept = request.headers.get("accept", SSE_CONTENT_TYPE)
        try:
            run_input = AGUIAdapter.build_run_input(await request.body())
        except ValidationError as e:
            return Response(
                content=json.dumps(e.json()),
                media_type="application/json",
                status_code=422,
            )

        # Prune large messages from history
        if hasattr(run_input, "messages"):
            run_input.messages = prune_large_messages(run_input.messages)

        adapter = AGUIAdapter(agent=agent, run_input=run_input, accept=accept)
        event_stream = adapter.run_stream()
        sse_stream = adapter.encode_stream(event_stream)

        return StreamingResponse(
            sse_stream,
            media_type=accept,
        )

    if enable_web_ui:
        web_ui = agent.to_web(instructions=SUPERVISOR_SYSTEM_PROMPT)
        app.mount("/", web_ui)
        logger.info(
            "Starting server on %s:%s (A2A at /a2a, AG-UI at /ag-ui, Web UI: %s)",
            host,
            port,
            "Enabled at /" if enable_web_ui else "Disabled",
        )

    uvicorn.run(
        app,
        host=host,
        port=port,
        timeout_keep_alive=1800,
        timeout_graceful_shutdown=60,
        log_level="debug" if debug else "info",
    )


def agent_server():
    print(f"microsoft_agent v{__version__}")
    parser = argparse.ArgumentParser(
        add_help=False, description=f"Run the {AGENT_NAME} A2A + AG-UI Server"
    )
    parser.add_argument(
        "--host", default=DEFAULT_HOST, help="Host to bind the server to"
    )
    parser.add_argument(
        "--port", type=int, default=DEFAULT_PORT, help="Port to bind the server to"
    )
    parser.add_argument("--debug", type=bool, default=DEFAULT_DEBUG, help="Debug mode")
    parser.add_argument("--reload", action="store_true", help="Enable auto-reload")

    parser.add_argument(
        "--provider",
        default=DEFAULT_PROVIDER,
        choices=["openai", "anthropic", "google", "huggingface"],
        help="LLM Provider",
    )
    parser.add_argument("--model-id", default=DEFAULT_MODEL_ID, help="LLM Model ID")
    parser.add_argument(
        "--base-url",
        default=DEFAULT_LLM_BASE_URL,
        help="LLM Base URL (for OpenAI compatible providers)",
    )
    parser.add_argument("--api-key", default=DEFAULT_LLM_API_KEY, help="LLM API Key")
    parser.add_argument("--mcp-url", default=DEFAULT_MCP_URL, help="MCP Server URL")
    parser.add_argument(
        "--mcp-config", default=DEFAULT_MCP_CONFIG, help="MCP Server Config"
    )
    parser.add_argument(
        "--web",
        action="store_true",
        default=DEFAULT_ENABLE_WEB_UI,
        help="Enable Pydantic AI Web UI",
    )
    parser.add_argument("--help", action="store_true", help="Show usage")

    args = parser.parse_args()

    if hasattr(args, "help") and args.help:

        usage()

        sys.exit(0)

    if args.debug:
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[logging.StreamHandler()],
            force=True,
        )
        logging.getLogger("pydantic_ai").setLevel(logging.DEBUG)
        logging.getLogger("fastmcp").setLevel(logging.DEBUG)
        logging.getLogger("httpcore").setLevel(logging.DEBUG)
        logging.getLogger("httpx").setLevel(logging.DEBUG)
        logger.setLevel(logging.DEBUG)
        logger.debug("Debug mode enabled")

    create_agent_server(
        provider=args.provider,
        model_id=args.model_id,
        base_url=args.base_url,
        api_key=args.api_key,
        mcp_url=args.mcp_url,
        mcp_config=args.mcp_config,
        debug=args.debug,
        host=args.host,
        port=args.port,
        enable_web_ui=args.web,
        ssl_verify=not args.insecure,
    )


def usage():
    print(
        f"Microsoft Agent ({__version__}): CLI Tool\n\n"
        "Usage:\n"
        "--host          [ Host to bind the server to ]\n"
        "--port          [ Port to bind the server to ]\n"
        "--debug         [ Debug mode ]\n"
        "--reload        [ Enable auto-reload ]\n"
        "--provider      [ LLM Provider ]\n"
        "--model-id      [ LLM Model ID ]\n"
        "--base-url      [ LLM Base URL (for OpenAI compatible providers) ]\n"
        "--api-key       [ LLM API Key ]\n"
        "--mcp-url       [ MCP Server URL ]\n"
        "--mcp-config    [ MCP Server Config ]\n"
        "--web           [ Enable Pydantic AI Web UI ]\n"
        "\n"
        "Examples:\n"
        "  [Simple]  microsoft-agent \n"
        '  [Complex] microsoft-agent --host "value" --port "value" --debug "value" --reload --provider "value" --model-id "value" --base-url "value" --api-key "value" --mcp-url "value" --mcp-config "value" --web\n'
    )


if __name__ == "__main__":
    agent_server()
