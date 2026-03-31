"""Microsoft Graph graph configuration — tag prompts and env var mappings.

This is the only file needed to enable graph mode for this agent.
Provides TAG_PROMPTS and TAG_ENV_VARS for create_graph_agent_server().
"""

                                                                       
TAG_PROMPTS: dict[str, str] = {
    "admin": (
        "You are a Microsoft Graph Admin specialist. Help users manage and interact with Admin functionality using the available tools."
    ),
    "agreements": (
        "You are a Microsoft Graph Agreements specialist. Help users manage and interact with Agreements functionality using the available tools."
    ),
    "applications": (
        "You are a Microsoft Graph Applications specialist. Help users manage and interact with Applications functionality using the available tools."
    ),
    "audit": (
        "You are a Microsoft Graph Audit specialist. Help users manage and interact with Audit functionality using the available tools."
    ),
    "auth": (
        "You are a Microsoft Graph Auth specialist. Help users manage and interact with Auth functionality using the available tools."
    ),
    "calendar": (
        "You are a Microsoft Graph Calendar specialist. Help users manage and interact with Calendar functionality using the available tools."
    ),
    "chat": (
        "You are a Microsoft Graph Chat specialist. Help users manage and interact with Chat functionality using the available tools."
    ),
    "communications": (
        "You are a Microsoft Graph Communications specialist. Help users manage and interact with Communications functionality using the available tools."
    ),
    "connections": (
        "You are a Microsoft Graph Connections specialist. Help users manage and interact with Connections functionality using the available tools."
    ),
    "contacts": (
        "You are a Microsoft Graph Contacts specialist. Help users manage and interact with Contacts functionality using the available tools."
    ),
    "devices": (
        "You are a Microsoft Graph Devices specialist. Help users manage and interact with Devices functionality using the available tools."
    ),
    "directory": (
        "You are a Microsoft Graph Directory specialist. Help users manage and interact with Directory functionality using the available tools."
    ),
    "domains": (
        "You are a Microsoft Graph Domains specialist. Help users manage and interact with Domains functionality using the available tools."
    ),
    "education": (
        "You are a Microsoft Graph Education specialist. Help users manage and interact with Education functionality using the available tools."
    ),
    "employee_experience": (
        "You are a Microsoft Graph Employee Experience specialist. Help users manage and interact with Employee Experience functionality using the available tools."
    ),
    "files": (
        "You are a Microsoft Graph Files specialist. Help users manage and interact with Files functionality using the available tools."
    ),
    "groups": (
        "You are a Microsoft Graph Groups specialist. Help users manage and interact with Groups functionality using the available tools."
    ),
    "identity": (
        "You are a Microsoft Graph Identity specialist. Help users manage and interact with Identity functionality using the available tools."
    ),
    "mail": (
        "You are a Microsoft Graph Mail specialist. Help users manage and interact with Mail functionality using the available tools."
    ),
    "meta": (
        "You are a Microsoft Graph Meta specialist. Help users manage and interact with Meta functionality using the available tools."
    ),
    "notes": (
        "You are a Microsoft Graph Notes specialist. Help users manage and interact with Notes functionality using the available tools."
    ),
    "organization": (
        "You are a Microsoft Graph Organization specialist. Help users manage and interact with Organization functionality using the available tools."
    ),
    "places": (
        "You are a Microsoft Graph Places specialist. Help users manage and interact with Places functionality using the available tools."
    ),
    "policies": (
        "You are a Microsoft Graph Policies specialist. Help users manage and interact with Policies functionality using the available tools."
    ),
    "print": (
        "You are a Microsoft Graph Print specialist. Help users manage and interact with Print functionality using the available tools."
    ),
    "privacy": (
        "You are a Microsoft Graph Privacy specialist. Help users manage and interact with Privacy functionality using the available tools."
    ),
    "reports": (
        "You are a Microsoft Graph Reports specialist. Help users manage and interact with Reports functionality using the available tools."
    ),
    "search": (
        "You are a Microsoft Graph Search specialist. Help users manage and interact with Search functionality using the available tools."
    ),
    "security": (
        "You are a Microsoft Graph Security specialist. Help users manage and interact with Security functionality using the available tools."
    ),
    "sites": (
        "You are a Microsoft Graph Sites specialist. Help users manage and interact with Sites functionality using the available tools."
    ),
    "solutions": (
        "You are a Microsoft Graph Solutions specialist. Help users manage and interact with Solutions functionality using the available tools."
    ),
    "storage": (
        "You are a Microsoft Graph Storage specialist. Help users manage and interact with Storage functionality using the available tools."
    ),
    "subscriptions": (
        "You are a Microsoft Graph Subscriptions specialist. Help users manage and interact with Subscriptions functionality using the available tools."
    ),
    "tasks": (
        "You are a Microsoft Graph Tasks specialist. Help users manage and interact with Tasks functionality using the available tools."
    ),
    "teams": (
        "You are a Microsoft Graph Teams specialist. Help users manage and interact with Teams functionality using the available tools."
    ),
    "user": (
        "You are a Microsoft Graph User specialist. Help users manage and interact with User functionality using the available tools."
    ),
}


                                                                        
TAG_ENV_VARS: dict[str, str] = {
    "admin": "ADMINTOOL",
    "agreements": "AGREEMENTSTOOL",
    "applications": "APPLICATIONSTOOL",
    "audit": "AUDITTOOL",
    "auth": "AUTHTOOL",
    "calendar": "CALENDARTOOL",
    "chat": "CHATTOOL",
    "communications": "COMMUNICATIONSTOOL",
    "connections": "CONNECTIONSTOOL",
    "contacts": "CONTACTSTOOL",
    "devices": "DEVICESTOOL",
    "directory": "DIRECTORYTOOL",
    "domains": "DOMAINSTOOL",
    "education": "EDUCATIONTOOL",
    "employee_experience": "EMPLOYEE_EXPERIENCETOOL",
    "files": "FILESTOOL",
    "groups": "GROUPSTOOL",
    "identity": "IDENTITYTOOL",
    "mail": "MAILTOOL",
    "meta": "METATOOL",
    "notes": "NOTESTOOL",
    "organization": "ORGANIZATIONTOOL",
    "places": "PLACESTOOL",
    "policies": "POLICIESTOOL",
    "print": "PRINTTOOL",
    "privacy": "PRIVACYTOOL",
    "reports": "REPORTSTOOL",
    "search": "SEARCHTOOL",
    "security": "SECURITYTOOL",
    "sites": "SITESTOOL",
    "solutions": "SOLUTIONSTOOL",
    "storage": "STORAGETOOL",
    "subscriptions": "SUBSCRIPTIONSTOOL",
    "tasks": "TASKSTOOL",
    "teams": "TEAMSTOOL",
    "user": "USERTOOL",
}
