# Microsoft Agent - A2A | AG-UI | MCP

![PyPI - Version](https://img.shields.io/pypi/v/microsoft-agent)
![MCP Server](https://badge.mcpx.dev?type=server 'MCP Server')
![PyPI - Downloads](https://img.shields.io/pypi/dd/microsoft-agent)
![GitHub Repo stars](https://img.shields.io/github/stars/Knuckles-Team/microsoft-agent)
![GitHub forks](https://img.shields.io/github/forks/Knuckles-Team/microsoft-agent)
![GitHub contributors](https://img.shields.io/github/contributors/Knuckles-Team/microsoft-agent)
![PyPI - License](https://img.shields.io/pypi/l/microsoft-agent)
![GitHub](https://img.shields.io/github/license/Knuckles-Team/microsoft-agent)

![GitHub last commit (by committer)](https://img.shields.io/github/last-commit/Knuckles-Team/microsoft-agent)
![GitHub pull requests](https://img.shields.io/github/issues-pr/Knuckles-Team/microsoft-agent)
![GitHub closed pull requests](https://img.shields.io/github/issues-pr-closed/Knuckles-Team/microsoft-agent)
![GitHub issues](https://img.shields.io/github/issues/Knuckles-Team/microsoft-agent)

![GitHub top language](https://img.shields.io/github/languages/top/Knuckles-Team/microsoft-agent)
![GitHub language count](https://img.shields.io/github/languages/count/Knuckles-Team/microsoft-agent)
![GitHub repo size](https://img.shields.io/github/repo-size/Knuckles-Team/microsoft-agent)
![GitHub repo file count (file type)](https://img.shields.io/github/directory-file-count/Knuckles-Team/microsoft-agent)
![PyPI - Wheel](https://img.shields.io/pypi/wheel/microsoft-agent)
![PyPI - Implementation](https://img.shields.io/pypi/implementation/microsoft-agent)

*Version: 0.1.2*

## Overview

Microsoft Graph MCP Server + A2A Supervisor Agent

It includes a Model Context Protocol (MCP) server that wraps the Microsoft Graph API and an out-of-the-box Agent2Agent (A2A) Supervisor Agent.

Manage your Microsoft 365 tenant (Users, Groups, Calendars, Drive, etc.) through natural language!

This repository is actively maintained - Contributions are welcome!

### Capabilities:
- **Comprehensive Graph API Coverage**: Access thousands of Microsoft Graph endpoints via MCP tools.
- **Supervisor-Worker Agent Architecture**: A smart supervisor delegates tasks to specialized agents (e.g., Calendar Agent, User Agent).
- **Secure Authentication**: Supports OAuth, OIDC, and other authentication methods.

## MCP

### MCP Tools

This server provides tools for a vast array of Microsoft Graph resources. Due to the large number of tools, they are organized by resource type.

**Supported Resources (Partial List):**
- **Users**: `get_user`, `update_user`, `list_user`, etc.
- **Groups**: `get_group`, `post_groups_group`, `list_members_group`, etc.
- **Calendar**: `get_calendar`, `post_events`, `list_calendarview`, etc.
- **Drive/DriveItems**: `get_drive`, `search_driveitem`, `upload_driveitem`, etc.
- **Mail**: `send_mail`, `list_messages`, etc.
- **Directory Objects**: `get_directoryobject`, `check_member_objects`, etc.
- **Planner**, **OneNote**, **Teams**, and more.

All tools generally follow the naming convention: `action_resource` (e.g., `list_user`, `delete_group`).

### Using as an MCP Server

The MCP Server can be run in two modes: `stdio` (for local testing) or `http` (for networked access).

#### Run in stdio mode (default):
```bash
microsoft-agent --transport "stdio"
```

#### Run in HTTP mode:
```bash
microsoft-agent --transport "http"  --host "0.0.0.0"  --port "8000"
```

AI Prompt:
```text
Find who manages the 'Engineering' group and list its members.
```

AI Response:
```text
I've found the 'Engineering' group.
Owners:
- Jane Doe (Head of Engineering)

Members:
- John Smith
- Alice Johnson
- Bob Williams
...
```

## A2A Agent

This package includes a powerful A2A Supervisor Agent that orchestrates interaction with the Microsoft MCP tools.

### Architecture

The system uses a Supervisor Agent that analyzes user requests and delegates them to domain-specific Child Agents.

```mermaid
---
config:
  layout: dagre
---
flowchart TB
 subgraph subGraph0["Agent System"]
        S["Supervisor Agent"]
        subGraph1["Specialized Agents"]
            CA["Calendar Agent"]
            GA["Group Agent"]
            UA["User Agent"]
            DA["Drive Agent"]
            OA["...Other Agents"]
        end
        B["A2A Server"]
        M["MCP Tools"]
  end
    U["User Query"] --> B
    B --> S
    S --Delegates--> CA & GA & UA & DA & OA
    CA & GA & UA & DA & OA --> M
    M --> api["Microsoft Graph API"]

     S:::agent
     B:::server
     U:::server
     CA:::worker
     GA:::worker
     UA:::worker
     DA:::worker
     OA:::worker

    classDef server fill:#f9f,stroke:#333
    classDef agent fill:#bbf,stroke:#333,stroke-width:2px
    classDef worker fill:#dcedc8,stroke:#333
    style B stroke:#000000,fill:#FFD600
    style M stroke:#000000,fill:#BBDEFB
    style U fill:#C8E6C9
    style subGraph0 fill:#FFF9C4
```

### Component Interaction

1. **User** sends a request (e.g., "Schedule a meeting with the Engineering team").
2. **Supervisor Agent** identifies this as a calendar and group task.
3. **Supervisor** delegates finding the group members to the **Group Agent**.
4. **Group Agent** calls `list_members_group` tool and returns emails.
5. **Supervisor** delegates scheduling to the **Calendar Agent** with the retrieved emails.
6. **Calendar Agent** calls `post_events` tool.
7. **Supervisor** confirms completion to the User.

## Usage

### MCP CLI

| Short Flag | Long Flag                          | Description                                                                 |
|------------|------------------------------------|-----------------------------------------------------------------------------|
| -h         | --help                             | Display help information                                                    |
| -t         | --transport                        | Transport method: 'stdio', 'http', or 'sse' [legacy] (default: stdio)       |
| -s         | --host                             | Host address for HTTP transport (default: 0.0.0.0)                          |
| -p         | --port                             | Port number for HTTP transport (default: 8000)                              |
|            | --auth-type                        | Auth type: 'none', 'static', 'jwt', 'oauth-proxy', 'oidc-proxy' (default: none) |
|            | ...                                | (See standard FastMCP auth flags)                                           |

### A2A CLI

#### Endpoints
- **Web UI**: `http://localhost:9000/` (if enabled)
- **A2A**: `http://localhost:9000/a2a` (Discovery: `/a2a/.well-known/agent.json`)
- **AG-UI**: `http://localhost:9000/ag-ui` (POST)

| Argument          | Description                                                    | Default                        |
|-------------------|----------------------------------------------------------------|--------------------------------|
| `--host`          | Host to bind the server to                                     | `0.0.0.0`                      |
| `--port`          | Port to bind the server to                                     | `9000`                         |
| `--provider`      | LLM Provider (openai, anthropic, google, huggingface)          | `openai`                       |
| `--model-id`      | LLM Model ID                                                   | `qwen/qwen3-4b-2507`           |
| `--mcp-url`       | MCP Server URL                                                 | `http://microsoft-agent:8000/mcp` |

### Examples

#### Run A2A Server
```bash
microsoft-agent-server --provider openai --model-id gpt-4o --api-key sk-... --mcp-url http://localhost:8000/mcp
```

## Docker

### Build

```bash
docker build -t microsoft-agent .
```

### Run MCP Server

```bash
docker run -p 8000:8000 microsoft-agent
```

### Run Agent Server

```bash
docker run -e CMD=agent-server -p 9000:9000 microsoft-agent
```

### Deploy as a Service

```bash
docker pull knucklessg1/microsoft-agent:latest

docker run -d \
  --name microsoft-agent \
  -p 8000:8000 \
  -e HOST=0.0.0.0 \
  -e PORT=8000 \
  -e TRANSPORT=http \
  knucklessg1/microsoft-agent:latest
```

## Install Python Package

```bash
python -m pip install microsoft-agent
```
```bash
uv pip install microsoft-agent
```

## Repository Owners

<img width="100%" height="180em" src="https://github-readme-stats.vercel.app/api?username=Knucklessg1&show_icons=true&hide_border=true&&count_private=true&include_all_commits=true" />

![GitHub followers](https://img.shields.io/github/followers/Knucklessg1)
![GitHub User's stars](https://img.shields.io/github/stars/Knucklessg1)

Documentation:

[Microsoft API Docs](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)