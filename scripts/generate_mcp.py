import json
import os
import re
import sys

CATEGORIES = {
    "mail": ["mail", "message", "attachment", "folder", "draft"],
    "calendar": ["calendar", "event", "meeting"],
    "files": [
        "drive",
        "file",
        "folder",
        "item",
        "workbook",
        "chart",
        "range",
        "list",
        "upload",
        "download",
    ],
    "notes": ["onenote", "notebook", "section", "page"],
    "tasks": ["todo", "task", "plan"],
    "contacts": ["contact"],
    "user": ["user", "me"],
    "chat": ["chat"],
    "teams": ["team", "channel"],
    "search": ["search", "query"],
    "sites": ["site"],
}


def get_tags(tool_name):
    tags = []
    lower_name = tool_name.lower()
    for cat, keywords in CATEGORIES.items():
        if any(kw in lower_name for kw in keywords):
            tags.append(cat)
    return tags


def generate_tool(endpoint):
    tool_name = endpoint["toolName"].replace("-", "_")
    path_pattern = endpoint["pathPattern"]
    method = endpoint["method"].upper()

    description = f"{tool_name}: {method} {path_pattern}"

    if "llmTip" in endpoint:
        tip = endpoint["llmTip"].replace('"', "'")
        description += f"\\n\\nTIP: {tip}"

    tags = get_tags(tool_name)
    tags_str = f", tags={json.dumps(tags)}" if tags else ""

    path_params = re.findall(r"\{([a-zA-Z0-9_-]+)\}", path_pattern)

    func_args = []
    path_args = []

    for param in path_params:
        py_param = param.replace("-", "_")
        func_args.append(
            f'{py_param}: str = Field(..., description="Parameter for {param}")'
        )
        path_args.append(f'"{param}": {py_param}')

    if method in ["POST", "PUT", "PATCH"]:
        func_args.append(
            'data: Optional[Dict[str, Any]] = Field(None, description="Request body data")'
        )

    func_args.append(
        'params: Optional[Dict[str, Any]] = Field(None, description="Query parameters")'
    )
    if "supportsTimezone" in endpoint:
        func_args.append(
            'timezone: Optional[str] = Field(None, description="IANA timezone")'
        )

    timezone_logic = ""
    if "supportsTimezone" in endpoint:
        timezone_logic = "if timezone: request_headers['Prefer'] = f'outlook.timezone=\"{timezone}\"'"

    args_str = ", ".join(func_args)
    path_args_str = ", ".join(path_args)

    func_def = f'''
    @mcp.tool(name="{tool_name}", description="""{description}"""{tags_str})
    def {tool_name}({args_str}) -> Any:
        """{description}"""
        client = get_client()
        path = "{path_pattern}"
        # Replace path parameters
        for k, v in {{{path_args_str}}}.items():
            path = path.replace(f"{{{{k}}}}", v)

        method = "{method}"
        request_params = params or {{}}
        request_headers = {{}}

        {timezone_logic}

        kwargs = {{"headers": request_headers}}
        if method in ["POST", "PUT", "PATCH"]:
             kwargs["json"] = data

        return client.graph_request(method, path, params=request_params, **kwargs)
'''
    return func_def


def main():
    endpoints_path = "/home/genius/Workspace/microsoft-agent/ms-365-mcp-server-main/src/endpoints.json"
    if not os.path.exists(endpoints_path):
        print(f"Error: {endpoints_path} not found", file=sys.stderr)
        return

    with open(endpoints_path, "r") as f:
        endpoints = json.load(f)

    print("#!/usr/bin/python")
    print("# coding: utf-8")
    print("")
    print("import os")
    print("import sys")
    print("from typing import Optional, List, Dict, Union, Any")
    print("from pydantic import Field")
    print("from fastmcp import FastMCP")
    print("from microsoft_agent.middlewares import get_client")
    print("from microsoft_agent.auth import AuthManager")
    print("from starlette.requests import Request")
    print("from starlette.responses import JSONResponse")
    print("")
    print("def register_tools(mcp: FastMCP):")
    print('    @mcp.custom_route("/health", methods=["GET"])')
    print("    async def health_check(request: Request) -> JSONResponse:")
    print('        return JSONResponse({"status": "OK"})')
    print("")
    print("    # Initialize AuthManager")
    print(
        '    CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")'
    )
    print('    AUTHORITY = "https://login.microsoftonline.com/common"')
    print(
        '    SCOPES = ["User.Read", "Mail.ReadWrite", "Calendars.ReadWrite", "Files.ReadWrite", "Tasks.ReadWrite", "Contacts.ReadWrite", "Group.ReadWrite.All", "Directory.Read.All", "Sites.Read.All", "Chat.Read", "ChatMessage.Read.All", "ChannelMessage.Read.All"]'
    )
    print("")
    print("    auth_manager = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)")
    print("")
    print(
        '    @mcp.tool(name="login", description="Authenticate with Microsoft using device code flow", tags=["auth"])'
    )
    print(
        '    def login(force: bool = Field(False, description="Force a new login even if already logged in")) -> str:'
    )
    print('        """Authenticate with Microsoft using device code flow"""')
    print("        if not force:")
    print("            token = auth_manager.get_token()")
    print("            if token:")
    print("                account = auth_manager.get_current_account()")
    print(
        '                username = account.get("username", "Unknown") if account else "Unknown"'
    )
    print(
        '                return f"Already logged in as {username}. Use force=True to login with a different account."'
    )
    print("")
    print("        def print_code(msg):")
    print('            print(f"\\n{msg}\\n")')
    print("        ")
    print("        try:")
    print("            return auth_manager.acquire_token_by_device_code(print_code)")
    print("        except Exception as e:")
    print('            return f"Authentication failed: {str(e)}"')
    print("")
    print(
        '    @mcp.tool(name="logout", description="Log out from Microsoft account", tags=["auth"])'
    )
    print("    def logout() -> str:")
    print('        """Log out from Microsoft account"""')
    print("        auth_manager.logout()")
    print('        return "Logged out successfully"')
    print("")
    print(
        '    @mcp.tool(name="verify_login", description="Check current Microsoft authentication status", tags=["auth"])'
    )
    print("    def verify_login() -> str:")
    print('        """Check current Microsoft authentication status"""')
    print("        token = auth_manager.get_token()")
    print("        if token:")
    print("            account = auth_manager.get_current_account()")
    print(
        '            username = account.get("username", "Unknown") if account else "Unknown"'
    )
    print('            return f"Logged in as {username}"')
    print('        return "Not logged in"')
    print("")
    print(
        '    @mcp.tool(name="list_accounts", description="List all available Microsoft accounts", tags=["auth"])'
    )
    print("    def list_accounts() -> str:")
    print('        """List all available Microsoft accounts"""')
    print("        accounts = auth_manager.list_accounts()")
    print("        if not accounts:")
    print('            return "No accounts found"')
    print("        ")
    print("        result = []")
    print("        current = auth_manager.get_current_account()")
    print('        current_id = current.get("home_account_id") if current else None')
    print("        ")
    print("        for acc in accounts:")
    print(
        '            is_selected = "*" if acc.get("home_account_id") == current_id else " "'
    )
    print(
        "            result.append(f\"{is_selected} {acc.get('username')} ({acc.get('name')})\")"
    )
    print("        ")
    print('        return "\\n".join(result)')
    print("")
    print(
        '    @mcp.tool(name="search_tools", description="Search available Microsoft Graph API tools", tags=["meta"])'
    )
    print(
        '    def search_tools(query: str = Field(..., description="Search query"), limit: int = Field(20, description="Max results")) -> str:'
    )
    print('        """Search available Microsoft Graph API tools"""')
    print("        import inspect")
    print("        import sys")
    print("        ")
    print("        results = []")
    print("        query = query.lower()")
    print(
        "        functions = inspect.getmembers(sys.modules[__name__], inspect.isfunction)"
    )
    print("        ")
    print("        for name, func in functions:")
    print(
        '            if "_" in name and not name.startswith("_") and name not in ["health_check", "register_tools", "to_boolean", "to_integer", "get_logger", "login", "logout", "verify_login", "list_accounts", "search_tools"]:'
    )
    print('                 doc = inspect.getdoc(func) or ""')
    print("                 if query in name.lower() or doc and query in doc.lower():")
    print(
        "                     results.append(f\"{name}: {doc.splitlines()[0] if doc else 'No description'}\")"
    )
    print("                     if len(results) >= limit:")
    print("                         break")
    print("        ")
    print("        if not results:")
    print('             return "No tools found matching query"')
    print("             ")
    print('        return "\\n".join(results)')

    for endpoint in endpoints:
        print(generate_tool(endpoint))


if __name__ == "__main__":
    main()
