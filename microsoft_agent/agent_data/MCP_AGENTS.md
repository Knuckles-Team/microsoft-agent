# MCP_AGENTS.md - Dynamic Agent Registry

This file tracks the generated agents from MCP servers. You can manually modify the 'Tools' list to customize agent expertise.

## Agent Mapping Table

| Name | Description | System Prompt | Tools | Tag | Source MCP |
|------|-------------|---------------|-------|-----|------------|
| Microsoft Auth Specialist | Expert specialist for auth domain tasks. | You are a Microsoft Auth specialist. Help users manage and interact with Auth functionality using the available tools. | microsoft-agent_auth_toolset | auth | microsoft-agent |
| Microsoft Groups Specialist | Expert specialist for groups domain tasks. | You are a Microsoft Groups specialist. Help users manage and interact with Groups functionality using the available tools. | microsoft-agent_groups_toolset | groups | microsoft-agent |
| Microsoft Agreements Specialist | Expert specialist for agreements domain tasks. | You are a Microsoft Agreements specialist. Help users manage and interact with Agreements functionality using the available tools. | microsoft-agent_agreements_toolset | agreements | microsoft-agent |
| Microsoft Files Specialist | Expert specialist for files domain tasks. | You are a Microsoft Files specialist. Help users manage and interact with Files functionality using the available tools. | microsoft-agent_files_toolset | files | microsoft-agent |
| Microsoft Notes Specialist | Expert specialist for notes domain tasks. | You are a Microsoft Notes specialist. Help users manage and interact with Notes functionality using the available tools. | microsoft-agent_notes_toolset | notes | microsoft-agent |
| Microsoft Organization Specialist | Expert specialist for organization domain tasks. | You are a Microsoft Organization specialist. Help users manage and interact with Organization functionality using the available tools. | microsoft-agent_organization_toolset | organization | microsoft-agent |
| Microsoft Audit Specialist | Expert specialist for audit domain tasks. | You are a Microsoft Audit specialist. Help users manage and interact with Audit functionality using the available tools. | microsoft-agent_audit_toolset | audit | microsoft-agent |
| Microsoft Places Specialist | Expert specialist for places domain tasks. | You are a Microsoft Places specialist. Help users manage and interact with Places functionality using the available tools. | microsoft-agent_places_toolset | places | microsoft-agent |
| Microsoft Print Specialist | Expert specialist for print domain tasks. | You are a Microsoft Print specialist. Help users manage and interact with Print functionality using the available tools. | microsoft-agent_print_toolset | print | microsoft-agent |
| Microsoft Tasks Specialist | Expert specialist for tasks domain tasks. | You are a Microsoft Tasks specialist. Help users manage and interact with Tasks functionality using the available tools. | microsoft-agent_tasks_toolset | tasks | microsoft-agent |
| Microsoft Search Specialist | Expert specialist for search domain tasks. | You are a Microsoft Search specialist. Help users manage and interact with Search functionality using the available tools. | microsoft-agent_search_toolset | search | microsoft-agent |
| Microsoft Employee Experience Specialist | Expert specialist for employee_experience domain tasks. | You are a Microsoft Employee Experience specialist. Help users manage and interact with Employee Experience functionality using the available tools. | microsoft-agent_employee_experience_toolset | employee_experience | microsoft-agent |
| Microsoft Meta Specialist | Expert specialist for meta domain tasks. | You are a Microsoft Meta specialist. Help users manage and interact with Meta functionality using the available tools. | microsoft-agent_meta_toolset | meta | microsoft-agent |
| Microsoft Chat Specialist | Expert specialist for chat domain tasks. | You are a Microsoft Chat specialist. Help users manage and interact with Chat functionality using the available tools. | microsoft-agent_chat_toolset | chat | microsoft-agent |
| Microsoft Sites Specialist | Expert specialist for sites domain tasks. | You are a Microsoft Sites specialist. Help users manage and interact with Sites functionality using the available tools. | microsoft-agent_sites_toolset | sites | microsoft-agent |
| Microsoft Misc Specialist | Expert specialist for misc domain tasks. | You are a Microsoft Misc specialist. Help users manage and interact with Misc functionality using the available tools. | microsoft-agent_misc_toolset | misc | microsoft-agent |
| Microsoft Directory Specialist | Expert specialist for directory domain tasks. | You are a Microsoft Directory specialist. Help users manage and interact with Directory functionality using the available tools. | microsoft-agent_directory_toolset | directory | microsoft-agent |
| Microsoft Policies Specialist | Expert specialist for policies domain tasks. | You are a Microsoft Policies specialist. Help users manage and interact with Policies functionality using the available tools. | microsoft-agent_policies_toolset | policies | microsoft-agent |
| Microsoft Admin Specialist | Expert specialist for admin domain tasks. | You are a Microsoft Admin specialist. Help users manage and interact with Admin functionality using the available tools. | microsoft-agent_admin_toolset | admin | microsoft-agent |
| Microsoft Teams Specialist | Expert specialist for teams domain tasks. | You are a Microsoft Teams specialist. Help users manage and interact with Teams functionality using the available tools. | microsoft-agent_teams_toolset | teams | microsoft-agent |
| Microsoft Applications Specialist | Expert specialist for applications domain tasks. | You are a Microsoft Applications specialist. Help users manage and interact with Applications functionality using the available tools. | microsoft-agent_applications_toolset | applications | microsoft-agent |
| Microsoft Calendar Specialist | Expert specialist for calendar domain tasks. | You are a Microsoft Calendar specialist. Help users manage and interact with Calendar functionality using the available tools. | microsoft-agent_calendar_toolset | calendar | microsoft-agent |
| Microsoft Reports Specialist | Expert specialist for reports domain tasks. | You are a Microsoft Reports specialist. Help users manage and interact with Reports functionality using the available tools. | microsoft-agent_reports_toolset | reports | microsoft-agent |
| Microsoft Privacy Specialist | Expert specialist for privacy domain tasks. | You are a Microsoft Privacy specialist. Help users manage and interact with Privacy functionality using the available tools. | microsoft-agent_privacy_toolset | privacy | microsoft-agent |
| Microsoft Solutions Specialist | Expert specialist for solutions domain tasks. | You are a Microsoft Solutions specialist. Help users manage and interact with Solutions functionality using the available tools. | microsoft-agent_solutions_toolset | solutions | microsoft-agent |
| Microsoft Subscriptions Specialist | Expert specialist for subscriptions domain tasks. | You are a Microsoft Subscriptions specialist. Help users manage and interact with Subscriptions functionality using the available tools. | microsoft-agent_subscriptions_toolset | subscriptions | microsoft-agent |
| Microsoft Domains Specialist | Expert specialist for domains domain tasks. | You are a Microsoft Domains specialist. Help users manage and interact with Domains functionality using the available tools. | microsoft-agent_domains_toolset | domains | microsoft-agent |
| Microsoft User Specialist | Expert specialist for user domain tasks. | You are a Microsoft User specialist. Help users manage and interact with User functionality using the available tools. | microsoft-agent_user_toolset | user | microsoft-agent |
| Microsoft Connections Specialist | Expert specialist for connections domain tasks. | You are a Microsoft Connections specialist. Help users manage and interact with Connections functionality using the available tools. | microsoft-agent_connections_toolset | connections | microsoft-agent |
| Microsoft Storage Specialist | Expert specialist for storage domain tasks. | You are a Microsoft Storage specialist. Help users manage and interact with Storage functionality using the available tools. | microsoft-agent_storage_toolset | storage | microsoft-agent |
| Microsoft Security Specialist | Expert specialist for security domain tasks. | You are a Microsoft Security specialist. Help users manage and interact with Security functionality using the available tools. | microsoft-agent_security_toolset | security | microsoft-agent |
| Microsoft Devices Specialist | Expert specialist for devices domain tasks. | You are a Microsoft Devices specialist. Help users manage and interact with Devices functionality using the available tools. | microsoft-agent_devices_toolset | devices | microsoft-agent |
| Microsoft Contacts Specialist | Expert specialist for contacts domain tasks. | You are a Microsoft Contacts specialist. Help users manage and interact with Contacts functionality using the available tools. | microsoft-agent_contacts_toolset | contacts | microsoft-agent |
| Microsoft Education Specialist | Expert specialist for education domain tasks. | You are a Microsoft Education specialist. Help users manage and interact with Education functionality using the available tools. | microsoft-agent_education_toolset | education | microsoft-agent |
| Microsoft Identity Specialist | Expert specialist for identity domain tasks. | You are a Microsoft Identity specialist. Help users manage and interact with Identity functionality using the available tools. | microsoft-agent_identity_toolset | identity | microsoft-agent |
| Microsoft Communications Specialist | Expert specialist for communications domain tasks. | You are a Microsoft Communications specialist. Help users manage and interact with Communications functionality using the available tools. | microsoft-agent_communications_toolset | communications | microsoft-agent |
| Microsoft Mail Specialist | Expert specialist for mail domain tasks. | You are a Microsoft Mail specialist. Help users manage and interact with Mail functionality using the available tools. | microsoft-agent_mail_toolset | mail | microsoft-agent |

## Tool Inventory Table

| Tool Name | Description | Tag | Source |
|-----------|-------------|-----|--------|
| microsoft-agent_auth_toolset | Static hint toolset for auth based on config env. | auth | microsoft-agent |
| microsoft-agent_groups_toolset | Static hint toolset for groups based on config env. | groups | microsoft-agent |
| microsoft-agent_agreements_toolset | Static hint toolset for agreements based on config env. | agreements | microsoft-agent |
| microsoft-agent_files_toolset | Static hint toolset for files based on config env. | files | microsoft-agent |
| microsoft-agent_notes_toolset | Static hint toolset for notes based on config env. | notes | microsoft-agent |
| microsoft-agent_organization_toolset | Static hint toolset for organization based on config env. | organization | microsoft-agent |
| microsoft-agent_audit_toolset | Static hint toolset for audit based on config env. | audit | microsoft-agent |
| microsoft-agent_places_toolset | Static hint toolset for places based on config env. | places | microsoft-agent |
| microsoft-agent_print_toolset | Static hint toolset for print based on config env. | print | microsoft-agent |
| microsoft-agent_tasks_toolset | Static hint toolset for tasks based on config env. | tasks | microsoft-agent |
| microsoft-agent_search_toolset | Static hint toolset for search based on config env. | search | microsoft-agent |
| microsoft-agent_employee_experience_toolset | Static hint toolset for employee_experience based on config env. | employee_experience | microsoft-agent |
| microsoft-agent_meta_toolset | Static hint toolset for meta based on config env. | meta | microsoft-agent |
| microsoft-agent_chat_toolset | Static hint toolset for chat based on config env. | chat | microsoft-agent |
| microsoft-agent_sites_toolset | Static hint toolset for sites based on config env. | sites | microsoft-agent |
| microsoft-agent_misc_toolset | Static hint toolset for misc based on config env. | misc | microsoft-agent |
| microsoft-agent_directory_toolset | Static hint toolset for directory based on config env. | directory | microsoft-agent |
| microsoft-agent_policies_toolset | Static hint toolset for policies based on config env. | policies | microsoft-agent |
| microsoft-agent_admin_toolset | Static hint toolset for admin based on config env. | admin | microsoft-agent |
| microsoft-agent_teams_toolset | Static hint toolset for teams based on config env. | teams | microsoft-agent |
| microsoft-agent_applications_toolset | Static hint toolset for applications based on config env. | applications | microsoft-agent |
| microsoft-agent_calendar_toolset | Static hint toolset for calendar based on config env. | calendar | microsoft-agent |
| microsoft-agent_reports_toolset | Static hint toolset for reports based on config env. | reports | microsoft-agent |
| microsoft-agent_privacy_toolset | Static hint toolset for privacy based on config env. | privacy | microsoft-agent |
| microsoft-agent_solutions_toolset | Static hint toolset for solutions based on config env. | solutions | microsoft-agent |
| microsoft-agent_subscriptions_toolset | Static hint toolset for subscriptions based on config env. | subscriptions | microsoft-agent |
| microsoft-agent_domains_toolset | Static hint toolset for domains based on config env. | domains | microsoft-agent |
| microsoft-agent_user_toolset | Static hint toolset for user based on config env. | user | microsoft-agent |
| microsoft-agent_connections_toolset | Static hint toolset for connections based on config env. | connections | microsoft-agent |
| microsoft-agent_storage_toolset | Static hint toolset for storage based on config env. | storage | microsoft-agent |
| microsoft-agent_security_toolset | Static hint toolset for security based on config env. | security | microsoft-agent |
| microsoft-agent_devices_toolset | Static hint toolset for devices based on config env. | devices | microsoft-agent |
| microsoft-agent_contacts_toolset | Static hint toolset for contacts based on config env. | contacts | microsoft-agent |
| microsoft-agent_education_toolset | Static hint toolset for education based on config env. | education | microsoft-agent |
| microsoft-agent_identity_toolset | Static hint toolset for identity based on config env. | identity | microsoft-agent |
| microsoft-agent_communications_toolset | Static hint toolset for communications based on config env. | communications | microsoft-agent |
| microsoft-agent_mail_toolset | Static hint toolset for mail based on config env. | mail | microsoft-agent |
