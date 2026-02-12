import os
import re
import glob
from dataclasses import dataclass, field
from typing import List, Optional

DOCS_DIR = os.path.abspath(
    "/home/genius/Workspace/microsoft-agent/microsoft-documentation"
)
OUTPUT_DIR = os.path.abspath("/home/genius/Workspace/microsoft-agent/microsoft_agent")


@dataclass
class ApiEndpoint:
    name: str
    resource: str
    method: str
    url: str
    description: str
    parameters: List[str] = field(default_factory=list)
    original_parameters: List[str] = field(default_factory=list)
    doc_file: str = ""


def sanitize_name(name: str) -> str:
    """Sanitize string to be valid python identifier"""
    name = (
        name.replace("-", "_").replace("|", "_or_").replace(":", "_").replace(".", "_")
    )
    name = re.sub(r"[^a-zA-Z0-9_]", "", name)
    if name[0].isdigit():
        name = "v_" + name
    return name


def derive_param_name(param_raw: str, url_segment: str) -> str:
    """Derive a better parameter name based on context"""
    sanitized = sanitize_name(param_raw)

    if sanitized == "id":
        if url_segment:
            context = url_segment.rstrip("s")
            return f"{context}_id"

    return sanitized


def parse_markdown(file_path: str) -> Optional[ApiEndpoint]:
    filename = os.path.basename(file_path)
    if not filename.startswith("api_"):
        return None

    base_name = filename.replace("api_", "").replace(".md", "")
    parts = base_name.split("-")

    if len(parts) < 2:
        return None

    resource = parts[0]
    operation = "-".join(parts[1:])
    name = f"{operation}_{resource}"
    name = sanitize_name(name)

    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    http_match = re.search(
        r"HTTP request.*?```(.*?)```", content, re.DOTALL | re.IGNORECASE
    )
    method = "GET"
    url = ""

    if http_match:
        http_block = http_match.group(1).strip()
        lines = http_block.split("\n")
        valid_lines = [
            line.strip()
            for line in lines
            if line.strip() and not line.strip().startswith("//")
        ]

        request_line = ""
        candidates = []
        for line in valid_lines:
            parts_line = line.split()
            if len(parts_line) >= 2 and parts_line[0] in [
                "GET",
                "POST",
                "PATCH",
                "DELETE",
                "PUT",
            ]:
                candidates.append(line)

        if candidates:
            complex_candidates = [c for c in candidates if "{" in c]
            if complex_candidates:
                request_line = complex_candidates[0]
            else:
                request_line = candidates[0]

        if request_line:
            parts_req = request_line.split(maxsplit=1)
            method = parts_req[0]
            url = parts_req[1]

    if "?" in url:
        url = url.split("?")[0]

    description = ""
    title_match = re.search(r"^#\s+(.+)$", content, re.MULTILINE)
    if title_match:
        description = title_match.group(1)

    parameters = []
    original_parameters = []

    segments = url.split("/")

    current_params = []

    for i, seg in enumerate(segments):
        match = re.search(r"\{(.*?)\}", seg)
        if match:
            raw_param = match.group(1)
            prev_seg = segments[i - 1] if i > 0 else ""

            derived = derive_param_name(raw_param, prev_seg)

            base_derived = derived
            counter = 2
            while derived in current_params:
                derived = f"{base_derived}_{counter}"
                counter += 1

            current_params.append(derived)
            parameters.append(derived)
            original_parameters.append(raw_param)

    return ApiEndpoint(
        name=name,
        resource=sanitize_name(resource),
        method=method,
        url=url,
        description=description,
        parameters=parameters,
        original_parameters=original_parameters,
        doc_file=filename,
    )


def generate_api_code(endpoints: List[ApiEndpoint]) -> str:
    code = [
        "#!/usr/bin/env python",
        "# coding: utf-8",
        "",
        "import json",
        "import requests",
        "from typing import Dict, List, Optional, Any",
        "from urllib.parse import urljoin",
        "",
        "class Api:",
        "    def __init__(self, base_url: str = 'https://graph.microsoft.com/v1.0', token: Optional[str] = None):",
        "        self.base_url = base_url",
        "        self.token = token",
        "        self._session = requests.Session()",
        "",
        "    def get_headers(self) -> Dict[str, str]:",
        "        headers = {'Content-Type': 'application/json'}",
        "        if self.token:",
        "            headers['Authorization'] = f'Bearer {self.token}'",
        "        return headers",
        "",
        "    def request(self, method: str, endpoint: str, data: Dict = None, params: Dict = None) -> Any:",
        "        url = urljoin(self.base_url, endpoint.lstrip('/')) if not endpoint.startswith('http') else endpoint",
        "        headers = self.get_headers()",
        "        response = self._session.request(method=method, url=url, headers=headers, json=data, params=params)",
        "        if response.status_code >= 400:",
        "             try:",
        "                 err_msg = response.json()",
        "             except:",
        "                 err_msg = response.text",
        "             raise Exception(f'Error {response.status_code}: {err_msg}')",
        "        if response.status_code == 204:",
        "            return {'status': 'success'}",
        "        try:",
        "            return response.json()",
        "        except:",
        "            return response.text",
        "",
    ]

    code.append(f"    # --- Auto-generated Methods ({len(endpoints)} endpoints) ---")

    for ep in endpoints:
        args = ["self"]
        for p in ep.parameters:
            args.append(f"{p}: str")

        if ep.method in ["POST", "PATCH", "PUT"]:
            args.append("data: Dict = None")

        args.append("params: Dict = None")

        code.append("")
        code.append(f"    def {ep.name}({', '.join(args)}) -> Any:")
        code.append(f'        """{ep.description}"""')

        target_url = ep.url

        replaced_url = target_url
        for i, original in enumerate(ep.original_parameters):
            sanitized = ep.parameters[i]

            replaced_url = replaced_url.replace(
                f"{{{original}}}", f"{{{sanitized}}}", 1
            )

        code.append(f'        endpoint = f"{replaced_url}"')

        pass_data = "data" if "data" in args else "None"
        code.append(
            f'        return self.request("{ep.method}", endpoint, data={pass_data}, params=params)'
        )

    return "\n".join(code)


def generate_mcp_code(endpoints: List[ApiEndpoint]) -> str:
    code = [
        "#!/usr/bin/env python",
        "# coding: utf-8",
        "",
        "import os",
        "from typing import Optional, Dict, Any",
        "from fastmcp import FastMCP",
        "from pydantic import Field",
        "from .microsoft_api import Api",
        "from .utils import to_boolean",
        "",
        'mcp = FastMCP("Microsoft Agent")',
        "",
        "def get_api_client() -> Api:",
        '    token = os.environ.get("MICROSOFT_TOKEN")',
        "    return Api(token=token)",
        "",
    ]

    for ep in endpoints:
        code.append("")
        code.append(f"    # Resource: {ep.resource}")
        code.append(f'@mcp.tool(name="{ep.name}", description="{ep.description}")')

        args = []
        for p in ep.parameters:
            args.append(f'{p}: str = Field(..., description="Parameter for {p}")')

        if ep.method in ["POST", "PATCH", "PUT"]:
            args.append(
                'data: Optional[Dict[str, Any]] = Field(None, description="Request body data")'
            )

        args.append(
            'params: Optional[Dict[str, Any]] = Field(None, description="Query parameters")'
        )

        code.append(f"def {ep.name}({', '.join(args)}) -> Any:")
        code.append(f'    """{ep.description}"""')
        code.append("    client = get_api_client()")

        call_args = []
        for p in ep.parameters:
            call_args.append(f"{p}={p}")
        if ep.method in ["POST", "PATCH", "PUT"]:
            call_args.append("data=data")
        call_args.append("params=params")

        code.append(f"    return client.{ep.name}({', '.join(call_args)})")

    return "\n".join(code)


def generate_agent_code(endpoints: List[ApiEndpoint]) -> str:
    resources = sorted(list(set(ep.resource for ep in endpoints)))

    code = [
        "#!/usr/bin/env python",
        "# coding: utf-8",
        "",
        "import os",
        "import logging",
        "import argparse",
        "from typing import Any, Optional",
        "import uvicorn",
        "from fastapi import FastAPI",
        "from pydantic_ai import Agent, ModelSettings, RunContext",
        "from .utils import create_model",
        "from .microsoft_mcp import mcp",
        "",
        "# Basic Logging",
        "logging.basicConfig(level=logging.INFO)",
        "logger = logging.getLogger(__name__)",
        "",
        'AGENT_NAME = "MicrosoftAgent"',
        "",
        'SUPERVISOR_SYSTEM_PROMPT = """',
        "You are the Microsoft Graph Supervisor Agent.",
        "Your goal is to assist the user by delegating tasks to specialized sub-agents based on the resource type.",
        "Available resources: " + ", ".join(resources),
        '"""',
        "",
    ]

    for res in resources:
        code.append(
            f'{res.upper()}_AGENT_PROMPT = "You are the {res.capitalize()} Agent. You manage {res} resources using the available tools."'
        )

    code.append("")
    code.append("def create_agent(provider: str = 'openai', model_id: str = 'gpt-4o'):")
    code.append("    model = create_model(provider, model_id, None, None)")
    code.append("    ")
    code.append("    child_agents = {}")
    code.append("    ")
    code.append("    # Define tools for each sub-agent dynamically")
    code.append("    import inspect")
    code.append("    from . import microsoft_mcp")
    code.append("    ")
    code.append("    resource_tools = {}")
    code.append("    for name, func in inspect.getmembers(microsoft_mcp):")
    code.append(
        "        if hasattr(func, '__call__') and getattr(func, '__module__', '').endswith('microsoft_mcp'):"
    )
    code.append("            parts = name.split('_')")
    code.append("            if len(parts) > 1:")
    code.append("                 res = parts[-1]")
    code.append(
        "                 if res not in resource_tools: resource_tools[res] = []"
    )
    code.append("                 resource_tools[res].append(func)")

    for res in resources:
        code.append(
            f"    agent_{res} = Agent(model=model, system_prompt={res.upper()}_AGENT_PROMPT)"
        )
        code.append(f"    if '{res}' in resource_tools:")
        code.append(f"        for t in resource_tools['{res}']:")
        code.append(f"            agent_{res}.tool(t)")
        code.append(f"    child_agents['{res}'] = agent_{res}")

    code.append("    user = os.environ.get('USER', 'user')")
    code.append(
        "    supervisor = Agent(model=model, system_prompt=SUPERVISOR_SYSTEM_PROMPT, deps_type=Any)"
    )

    for res in resources:
        code.append("    @supervisor.tool")
        code.append(
            f"    async def delegate_to_{res}_agent(ctx: RunContext[Any], task: str) -> str:"
        )
        code.append(f'        """Delegate task to the {res.capitalize()} Agent."""')
        code.append(
            f"        result = await child_agents['{res}'].run(task, usage=ctx.usage)"
        )
        code.append("        return result.data")

    code.append("    return supervisor")

    code.append("")
    code.append("def agent_server():")
    code.append("    # Adapted from ansible-tower-mcp structure")
    code.append(
        "    parser = argparse.ArgumentParser(description='Microsoft Agent Server')"
    )
    code.append(
        "    parser.add_argument('--provider', type=str, default='openai', help='Model provider')"
    )
    code.append(
        "    parser.add_argument('--model', type=str, default='gpt-4o', help='Model ID')"
    )
    code.append(
        "    parser.add_argument('--port', type=int, default=8000, help='Port to run the server on')"
    )
    code.append("    args = parser.parse_args()")
    code.append("    ")
    code.append("    print(f'Starting {AGENT_NAME}...')")
    code.append("    agent = create_agent(provider=args.provider, model_id=args.model)")
    code.append("    print('Agent created successfully.')")
    code.append(
        "    # In a real run, we would start uvicorn here or interact with the agent."
    )

    return "\n".join(code)


def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    md_files = glob.glob(os.path.join(DOCS_DIR, "*.md"))
    endpoints = []

    print(f"Found {len(md_files)} documentation files.")

    for md in md_files:
        endpoint = parse_markdown(md)
        if endpoint:
            endpoints.append(endpoint)

    print(f"Successfully parsed {len(endpoints)} endpoints.")

    api_code = generate_api_code(endpoints)
    with open(os.path.join(OUTPUT_DIR, "microsoft_api.py"), "w") as f:
        f.write(api_code)
    print("Generated microsoft_api.py")

    mcp_code = generate_mcp_code(endpoints)
    with open(os.path.join(OUTPUT_DIR, "microsoft_mcp.py"), "w") as f:
        f.write(mcp_code)
    print("Generated microsoft_mcp.py")

    agent_code = generate_agent_code(endpoints)
    with open(os.path.join(OUTPUT_DIR, "microsoft_agent.py"), "w") as f:
        f.write(agent_code)
    print("Generated microsoft_agent.py")


if __name__ == "__main__":
    main()
