#!/usr/bin/python
import sys

# coding: utf-8
import json
import os
import argparse
import logging
import inspect
import uvicorn
import httpx
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

__version__ = "0.2.13"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
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
DEFAULT_CUSTOM_SKILLS_DIRECTORY = os.getenv("CUSTOM_SKILLS_DIRECTORY", None)
DEFAULT_ENABLE_WEB_UI = to_boolean(os.getenv("ENABLE_WEB_UI", "False"))
DEFAULT_SSL_VERIFY = to_boolean(os.getenv("SSL_VERIFY", "True"))

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


# =========================================================================
# Agent Prompts
# =========================================================================
ADMIN_AGENT_PROMPT = "You are the Admin Agent. You manage Microsoft 365 tenant administration including service health monitoring, service announcements, update messages, SharePoint admin settings, and delegated admin relationships."
AGREEMENTS_AGENT_PROMPT = "You are the Agreements Agent. You manage terms-of-use agreements using the available tools."
APPLICATIONS_AGENT_PROMPT = "You are the Applications Agent. You manage app registrations, service principals, credentials, and enterprise apps using the available tools."
AUDIT_AGENT_PROMPT = "You are the Audit Agent. You access directory audit logs, sign-in logs, and provisioning logs using the available tools."
AUTH_AGENT_PROMPT = "You are the Auth Agent. You manage authentication operations including login, logout, session verification, and account listing."
CALENDAR_AGENT_PROMPT = "You are the Calendar Agent. You manage calendar events, calendars, and scheduling using the available tools."
CHAT_AGENT_PROMPT = "You are the Chat Agent. You manage chats, chat messages, replies, and group conversations using the available tools."
COMMUNICATIONS_AGENT_PROMPT = "You are the Communications Agent. You manage online meetings (create, update, delete, list), call records, and user presence information using the available tools."
CONNECTIONS_AGENT_PROMPT = "You are the Connections Agent. You manage Microsoft Search external connections using the available tools."
CONTACTS_AGENT_PROMPT = "You are the Contacts Agent. You manage Outlook contacts (create, read, update, delete) using the available tools."
DEVICES_AGENT_PROMPT = "You are the Devices Agent. You manage directory devices, Intune managed devices, compliance policies, device configurations, and device actions like wipe and retire using the available tools."
DIRECTORY_AGENT_PROMPT = "You are the Directory Agent. You manage directory objects, roles, role templates, deleted items, role definitions, and role assignments using the available tools."
DOMAINS_AGENT_PROMPT = "You are the Domains Agent. You manage tenant domains including adding, verifying, deleting domains, and viewing DNS configuration records."
EDUCATION_AGENT_PROMPT = "You are the Education Agent. You manage education classes, schools, users, and assignments using the available tools."
EMPLOYEE_EXPERIENCE_AGENT_PROMPT = "You are the Employee Experience Agent. You manage learning providers and course activities using the available tools."
FILES_AGENT_PROMPT = "You are the Files Agent. You manage OneDrive files, Excel workbooks, OneNote notebooks, and SharePoint file operations using the available tools."
GROUPS_AGENT_PROMPT = "You are the Groups Agent. You manage Microsoft 365 groups, security groups, membership, ownership, conversations, and group drives using the available tools."
IDENTITY_AGENT_PROMPT = "You are the Identity Agent. You manage identity operations including guest user invitations, conditional access policies, access reviews, entitlement access packages, and lifecycle workflows."
MAIL_AGENT_PROMPT = "You are the Mail Agent. You manage email messages, folders, attachments, shared mailboxes, drafts, and mail operations using the available tools."
META_AGENT_PROMPT = "You are the Meta Agent. You provide tool discovery and search capabilities using the available tools."
NOTES_AGENT_PROMPT = "You are the Notes Agent. You manage OneNote notebooks, sections, and pages using the available tools."
ORGANIZATION_AGENT_PROMPT = "You are the Organization Agent. You manage organization profile, branding, and configuration using the available tools."
PLACES_AGENT_PROMPT = "You are the Places Agent. You manage rooms, room lists, and places using the available tools."
POLICIES_AGENT_PROMPT = "You are the Policies Agent. You manage authorization policies, token policies, permission grant policies, and admin consent policies using the available tools."
PRINT_AGENT_PROMPT = "You are the Print Agent. You manage printers, print jobs, and print shares using the available tools."
PRIVACY_AGENT_PROMPT = "You are the Privacy Agent. You manage subject rights requests for GDPR/CCPA compliance using the available tools."
REPORTS_AGENT_PROMPT = "You are the Reports Agent. You generate usage and activity reports for email, mailbox, Office 365, SharePoint, Teams, and OneDrive using the available tools."
SEARCH_AGENT_PROMPT = "You are the Search Agent. You execute Microsoft Graph search queries using the available tools."
SECURITY_AGENT_PROMPT = "You are the Security Agent. You manage security alerts, incidents, secure scores, threat intelligence, advanced hunting, risky users, risk detections, and sensitivity labels using the available tools."
SITES_AGENT_PROMPT = "You are the Sites Agent. You manage SharePoint sites, site lists, site drives, site items, and site administration using the available tools."
SOLUTIONS_AGENT_PROMPT = "You are the Solutions Agent. You manage booking businesses, appointments, and virtual events using the available tools."
STORAGE_AGENT_PROMPT = "You are the Storage Agent. You manage file storage containers using the available tools."
SUBSCRIPTIONS_AGENT_PROMPT = "You are the Subscriptions Agent. You manage webhook subscriptions (create, read, update, delete) using the available tools."
TASKS_AGENT_PROMPT = "You are the Tasks Agent. You manage Planner tasks, To-Do task lists, and task operations using the available tools."
TEAMS_AGENT_PROMPT = "You are the Teams Agent. You manage teams, channels, channel messages, and team membership using the available tools."
USER_AGENT_PROMPT = "You are the User Agent. You manage user profiles, user mail operations, meetings, and group membership using the available tools."


def create_agent(
    provider: str = DEFAULT_PROVIDER,
    model_id: str = DEFAULT_MODEL_ID,
    base_url: Optional[str] = DEFAULT_LLM_BASE_URL,
    api_key: Optional[str] = DEFAULT_LLM_API_KEY,
    mcp_url: str = DEFAULT_MCP_URL,
    mcp_config: str = DEFAULT_MCP_CONFIG,
    custom_skills_directory: Optional[str] = DEFAULT_CUSTOM_SKILLS_DIRECTORY,
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
        timeout=DEFAULT_TIMEOUT,
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

    agent_toolsets = []
    if mcp_url:
        if "sse" in mcp_url.lower():
            server = MCPServerSSE(
                mcp_url,
                http_client=httpx.AsyncClient(
                    verify=ssl_verify, timeout=DEFAULT_TIMEOUT
                ),
            )
        else:
            server = MCPServerStreamableHTTP(
                mcp_url,
                http_client=httpx.AsyncClient(
                    verify=ssl_verify, timeout=DEFAULT_TIMEOUT
                ),
            )
        agent_toolsets.append(server)
        logger.info(f"Connected to MCP Server: {mcp_url}")
    elif mcp_config:
        mcp_toolset = load_mcp_servers(mcp_config)
        for server in mcp_toolset:
            if hasattr(server, "http_client"):
                server.http_client = httpx.AsyncClient(
                    verify=ssl_verify, timeout=DEFAULT_TIMEOUT
                )
        agent_toolsets.extend(mcp_toolset)
        logger.info(f"Connected to MCP Config JSON: {mcp_toolset}")

    # Skills are loaded per-agent based on tags

    resource_tools: Dict[str, List] = {}
    for name, func in inspect.getmembers(microsoft_mcp):
        if hasattr(func, "__call__") and getattr(func, "__module__", "").endswith(
            "microsoft_mcp"
        ):
            parts = name.split("_")
            if len(parts) > 1:
                res = parts[-1]
                if res not in resource_tools:
                    resource_tools[res] = []
                resource_tools[res].append(func)

    # =========================================================================
    # Agent Definitions
    # =========================================================================
    agent_defs = {
        "admin": (ADMIN_AGENT_PROMPT, "Microsoft_Admin_Agent"),
        "agreements": (AGREEMENTS_AGENT_PROMPT, "Microsoft_Agreements_Agent"),
        "applications": (APPLICATIONS_AGENT_PROMPT, "Microsoft_Applications_Agent"),
        "audit": (AUDIT_AGENT_PROMPT, "Microsoft_Audit_Agent"),
        "auth": (AUTH_AGENT_PROMPT, "Microsoft_Auth_Agent"),
        "calendar": (CALENDAR_AGENT_PROMPT, "Microsoft_Calendar_Agent"),
        "chat": (CHAT_AGENT_PROMPT, "Microsoft_Chat_Agent"),
        "communications": (
            COMMUNICATIONS_AGENT_PROMPT,
            "Microsoft_Communications_Agent",
        ),
        "connections": (CONNECTIONS_AGENT_PROMPT, "Microsoft_Connections_Agent"),
        "contacts": (CONTACTS_AGENT_PROMPT, "Microsoft_Contacts_Agent"),
        "devices": (DEVICES_AGENT_PROMPT, "Microsoft_Devices_Agent"),
        "directory": (DIRECTORY_AGENT_PROMPT, "Microsoft_Directory_Agent"),
        "domains": (DOMAINS_AGENT_PROMPT, "Microsoft_Domains_Agent"),
        "education": (EDUCATION_AGENT_PROMPT, "Microsoft_Education_Agent"),
        "employee_experience": (
            EMPLOYEE_EXPERIENCE_AGENT_PROMPT,
            "Microsoft_Employee_Experience_Agent",
        ),
        "files": (FILES_AGENT_PROMPT, "Microsoft_Files_Agent"),
        "groups": (GROUPS_AGENT_PROMPT, "Microsoft_Groups_Agent"),
        "identity": (IDENTITY_AGENT_PROMPT, "Microsoft_Identity_Agent"),
        "mail": (MAIL_AGENT_PROMPT, "Microsoft_Mail_Agent"),
        "meta": (META_AGENT_PROMPT, "Microsoft_Meta_Agent"),
        "notes": (NOTES_AGENT_PROMPT, "Microsoft_Notes_Agent"),
        "organization": (ORGANIZATION_AGENT_PROMPT, "Microsoft_Organization_Agent"),
        "places": (PLACES_AGENT_PROMPT, "Microsoft_Places_Agent"),
        "policies": (POLICIES_AGENT_PROMPT, "Microsoft_Policies_Agent"),
        "print": (PRINT_AGENT_PROMPT, "Microsoft_Print_Agent"),
        "privacy": (PRIVACY_AGENT_PROMPT, "Microsoft_Privacy_Agent"),
        "reports": (REPORTS_AGENT_PROMPT, "Microsoft_Reports_Agent"),
        "search": (SEARCH_AGENT_PROMPT, "Microsoft_Search_Agent"),
        "security": (SECURITY_AGENT_PROMPT, "Microsoft_Security_Agent"),
        "sites": (SITES_AGENT_PROMPT, "Microsoft_Sites_Agent"),
        "solutions": (SOLUTIONS_AGENT_PROMPT, "Microsoft_Solutions_Agent"),
        "storage": (STORAGE_AGENT_PROMPT, "Microsoft_Storage_Agent"),
        "subscriptions": (SUBSCRIPTIONS_AGENT_PROMPT, "Microsoft_Subscriptions_Agent"),
        "tasks": (TASKS_AGENT_PROMPT, "Microsoft_Tasks_Agent"),
        "teams": (TEAMS_AGENT_PROMPT, "Microsoft_Teams_Agent"),
        "user": (USER_AGENT_PROMPT, "Microsoft_User_Agent"),
    }

    child_agents = {}

    for tag, (system_prompt, agent_name) in agent_defs.items():
        tag_toolsets = []
        for ts in agent_toolsets:

            def filter_func(ctx, tool_def, t=tag):
                return tool_in_tag(tool_def, t)

            if hasattr(ts, "filtered"):
                filtered_ts = ts.filtered(filter_func)
                tag_toolsets.append(filtered_ts)
            else:
                pass

        # Load specific skills for this tag
        skill_dir_name = f"microsoft-{tag}"

        # Check custom skills directory
        if custom_skills_directory:
            skill_dir_path = os.path.join(custom_skills_directory, skill_dir_name)
            if os.path.exists(skill_dir_path):
                tag_toolsets.append(SkillsToolset(directories=[skill_dir_path]))
                logger.info(
                    f"Loaded specialized skills for {tag} from {skill_dir_path}"
                )

        # Check default skills directory
        default_skill_path = os.path.join(get_skills_path(), skill_dir_name)
        if os.path.exists(default_skill_path):
            tag_toolsets.append(SkillsToolset(directories=[default_skill_path]))
            logger.info(
                f"Loaded specialized skills for {tag} from {default_skill_path}"
            )

        local_tools = resource_tools.get(tag, [])

        agent = Agent(
            model=model,
            system_prompt=system_prompt,
            name=agent_name,
            toolsets=tag_toolsets,
            tool_timeout=DEFAULT_TOOL_TIMEOUT,
            model_settings=settings,
        )

        for t in local_tools:
            agent.tool(t)

        child_agents[tag] = agent

    supervisor = Agent(
        model=model,
        system_prompt=SUPERVISOR_SYSTEM_PROMPT,
        model_settings=settings,
        name=AGENT_NAME,
        deps_type=Any,
    )

    # =========================================================================
    # Supervisor routing tools
    # =========================================================================

    @supervisor.tool
    async def assign_task_to_admin_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to tenant administration (service health, announcements, SharePoint admin, delegated admin) to the Admin Agent."""
        return (
            await child_agents["admin"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_agreements_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to terms-of-use agreements to the Agreements Agent."""
        return (
            await child_agents["agreements"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_applications_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to app registrations, service principals, or enterprise apps to the Applications Agent."""
        return (
            await child_agents["applications"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_audit_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to directory audits, sign-in logs, or provisioning logs to the Audit Agent."""
        return (
            await child_agents["audit"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_auth_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to authentication (login, logout, session verification) to the Auth Agent."""
        return (
            await child_agents["auth"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_calendar_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to calendar events, calendars, or scheduling to the Calendar Agent."""
        return (
            await child_agents["calendar"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_chat_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to chats, chat messages, replies, or group conversations to the Chat Agent."""
        return (
            await child_agents["chat"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_communications_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to online meetings, call records, or user presence to the Communications Agent."""
        return (
            await child_agents["communications"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_connections_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to Microsoft Search external connections to the Connections Agent."""
        return (
            await child_agents["connections"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_contacts_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to Outlook contacts management to the Contacts Agent."""
        return (
            await child_agents["contacts"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_devices_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to devices, managed devices, compliance policies, or device actions (wipe/retire) to the Devices Agent."""
        return (
            await child_agents["devices"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_directory_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to directory objects, roles, deleted items, role definitions, or role assignments to the Directory Agent."""
        return (
            await child_agents["directory"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_domains_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to tenant domain management (add, verify, delete, DNS records) to the Domains Agent."""
        return (
            await child_agents["domains"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_education_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to education classes, schools, users, or assignments to the Education Agent."""
        return (
            await child_agents["education"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_employee_experience_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to learning providers or course activities to the Employee Experience Agent."""
        return (
            await child_agents["employee_experience"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_files_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to OneDrive files, Excel workbooks, OneNote, or SharePoint file operations to the Files Agent."""
        return (
            await child_agents["files"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_groups_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to Microsoft 365 groups management (CRUD, membership, ownership, conversations, drives) to the Groups Agent."""
        return (
            await child_agents["groups"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_identity_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to identity operations (invitations, conditional access, access reviews, entitlements, lifecycle workflows) to the Identity Agent."""
        return (
            await child_agents["identity"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_mail_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to email messages, folders, attachments, shared mailboxes, or mail operations to the Mail Agent."""
        return (
            await child_agents["mail"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_meta_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to tool discovery or search to the Meta Agent."""
        return (
            await child_agents["meta"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_notes_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to OneNote notebooks, sections, or pages to the Notes Agent."""
        return (
            await child_agents["notes"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_organization_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to organization profile, branding, and configuration to the Organization Agent."""
        return (
            await child_agents["organization"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_places_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to rooms, room lists, or places management to the Places Agent."""
        return (
            await child_agents["places"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_policies_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to authorization, token, permission grant, or admin consent policies to the Policies Agent."""
        return (
            await child_agents["policies"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_print_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to printers, print jobs, or print shares to the Print Agent."""
        return (
            await child_agents["print"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_privacy_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to subject rights requests (GDPR/CCPA) to the Privacy Agent."""
        return (
            await child_agents["privacy"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_reports_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to usage and activity reports (email, mailbox, SharePoint, Teams, OneDrive) to the Reports Agent."""
        return (
            await child_agents["reports"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_search_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to Microsoft Graph search queries to the Search Agent."""
        return (
            await child_agents["search"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_security_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to security alerts, incidents, secure scores, threat intelligence, risky users, or sensitivity labels to the Security Agent."""
        return (
            await child_agents["security"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_sites_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to SharePoint sites, site lists, site drives, or site items to the Sites Agent."""
        return (
            await child_agents["sites"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_solutions_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to bookings, appointments, or virtual events to the Solutions Agent."""
        return (
            await child_agents["solutions"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_storage_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to file storage containers to the Storage Agent."""
        return (
            await child_agents["storage"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_subscriptions_agent(
        ctx: RunContext[Any], task: str
    ) -> str:
        """Assign a task related to webhook subscriptions to the Subscriptions Agent."""
        return (
            await child_agents["subscriptions"].run(
                task, usage=ctx.usage, deps=ctx.deps
            )
        ).output

    @supervisor.tool
    async def assign_task_to_tasks_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to Planner tasks, To-Do task lists, or task operations to the Tasks Agent."""
        return (
            await child_agents["tasks"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_teams_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to teams, channels, channel messages, or team membership to the Teams Agent."""
        return (
            await child_agents["teams"].run(task, usage=ctx.usage, deps=ctx.deps)
        ).output

    @supervisor.tool
    async def assign_task_to_user_agent(ctx: RunContext[Any], task: str) -> str:
        """Assign a task related to user profiles, mail operations, meetings, or group membership to the User Agent."""
        return (
            await child_agents["user"].run(task, usage=ctx.usage, deps=ctx.deps)
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
    base_url: Optional[str] = DEFAULT_LLM_BASE_URL,
    api_key: Optional[str] = DEFAULT_LLM_API_KEY,
    mcp_url: str = DEFAULT_MCP_URL,
    mcp_config: str = DEFAULT_MCP_CONFIG,
    custom_skills_directory: Optional[str] = DEFAULT_CUSTOM_SKILLS_DIRECTORY,
    debug: Optional[bool] = DEFAULT_DEBUG,
    host: Optional[str] = DEFAULT_HOST,
    port: Optional[int] = DEFAULT_PORT,
    enable_web_ui: bool = DEFAULT_ENABLE_WEB_UI,
    ssl_verify: bool = DEFAULT_SSL_VERIFY,
):
    print(
        f"Starting {AGENT_NAME}:"
        f"\tprovider={provider}"
        f"\tmodel={model_id}"
        f"\tbase_url={base_url}"
        f"\tmcp={mcp_url} | {mcp_config}"
        f"\tssl_verify={ssl_verify}"
    )
    agent = create_agent(
        provider=provider,
        model_id=model_id,
        base_url=base_url,
        api_key=api_key,
        mcp_url=mcp_url,
        mcp_config=mcp_config,
        custom_skills_directory=custom_skills_directory,
        ssl_verify=ssl_verify,
    )

    # Always load default skills

    skills = load_skills_from_directory(get_skills_path())

    logger.info(f"Loaded {len(skills)} default skills from {get_skills_path()}")

    # Load custom skills if provided

    if custom_skills_directory and os.path.exists(custom_skills_directory):

        custom_skills = load_skills_from_directory(custom_skills_directory)

        skills.extend(custom_skills)

        logger.info(
            f"Loaded {len(custom_skills)} custom skills from {custom_skills_directory}"
        )

    if not skills:

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
        "--custom-skills-directory",
        default=DEFAULT_CUSTOM_SKILLS_DIRECTORY,
        help="Directory containing additional custom agent skills",
    )
    parser.add_argument(
        "--web",
        action="store_true",
        default=DEFAULT_ENABLE_WEB_UI,
        help="Enable Pydantic AI Web UI",
    )

    parser.add_argument(
        "--insecure",
        action="store_true",
        help="Disable SSL verification for LLM requests (Use with caution)",
    )
    parser.add_argument("--help", action="store_true", help="Show usage")

    args = parser.parse_args()

    if hasattr(args, "help") and args.help:

        parser.print_help()

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
        custom_skills_directory=args.custom_skills_directory,
        debug=args.debug,
        host=args.host,
        port=args.port,
        enable_web_ui=args.web,
        ssl_verify=not args.insecure,
    )


if __name__ == "__main__":
    agent_server()
