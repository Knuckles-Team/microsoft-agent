---
name: microsoft-identity-directory
skill_type: skill
description: >-
  Entra ID directory: users, groups and membership, applications, service
  principals, domains, roles and assignments, conditional access, access
  reviews, entitlement, policies, organization settings and deleted items. Use
  when the agent must look up or administer identities and access. Do NOT use
  for security alerts (use microsoft-security-operations) or device management
  (use microsoft-endpoint-management).
license: MIT
tags: [microsoft, entra, directory, identity, graph]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft identity and directory administration

Typed directory operations. Role assignments, conditional access and application credentials are destructive-tier: require explicit approval.

## When to use
- Look up users, groups, members and owners.
- Create or change a group; add or remove members.
- Register and maintain applications and service principals.
- Review role definitions and assignments, conditional access and policies.
- Verify a domain; restore a deleted item.

## When NOT to use
- Alerts, incidents and sign-in risk: use microsoft-security-operations.
- Managed devices: use microsoft-endpoint-management.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (73)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_add_application_password` | W | app_id, data, params |
| `microsoft_add_group_member` | W | data, group_id, params |
| `microsoft_create_agreement` | W | data, params |
| `microsoft_create_application` | W | data, params |
| `microsoft_create_conditional_access_policy` | W | data, params |
| `microsoft_create_domain` | W | data, params |
| `microsoft_create_group` | W | data, params |
| `microsoft_create_invitation` | W | data, params |
| `microsoft_create_role_assignment` | W | data, params |
| `microsoft_create_service_principal` | W | data, params |
| `microsoft_delete_agreement` | D | agreement_id, confirm, params |
| `microsoft_delete_application` | D | app_id, confirm, params |
| `microsoft_delete_conditional_access_policy` | D | confirm, params, policy_id |
| `microsoft_delete_device` | D | confirm, device_id, params |
| `microsoft_delete_domain` | D | confirm, domain_id, params |
| `microsoft_delete_group` | D | confirm, group_id, params |
| `microsoft_delete_service_principal` | D | confirm, params, sp_id |
| `microsoft_get_access_review` | R | params, review_id |
| `microsoft_get_admin_consent_policy` | R | params |
| `microsoft_get_agreement` | R | agreement_id, params |
| `microsoft_get_application` | R | app_id, params |
| `microsoft_get_authorization_policy` | R | params |
| `microsoft_get_conditional_access_policy` | R | params, policy_id |
| `microsoft_get_delegated_admin_relationship` | R | params, rel_id |
| `microsoft_get_device` | R | device_id, params |
| `microsoft_get_directory_object` | R | object_id, params |
| `microsoft_get_directory_role` | R | params, role_id |
| `microsoft_get_domain` | R | domain_id, params |
| `microsoft_get_group` | R | group_id, params |
| `microsoft_get_me` | R | params |
| `microsoft_get_org_branding` | R | org_id, params |
| `microsoft_get_organization` | R | org_id, params |
| `microsoft_get_role_assignment` | R | assignment_id, params |
| `microsoft_get_role_definition` | R | params, role_id |
| `microsoft_get_service_principal` | R | params, sp_id |
| `microsoft_list_access_reviews` | R | params |
| `microsoft_list_agreements` | R | params |
| `microsoft_list_applications` | R | params |
| `microsoft_list_conditional_access_policies` | R | params |
| `microsoft_list_delegated_admin_relationships` | R | params |
| `microsoft_list_deleted_items` | R | params |
| `microsoft_list_devices` | R | params |
| `microsoft_list_directory_objects` | R | params |
| `microsoft_list_directory_role_templates` | R | params |
| `microsoft_list_directory_roles` | R | params |
| `microsoft_list_domain_service_configuration_records` | R | domain_id, params |
| `microsoft_list_domains` | R | params |
| `microsoft_list_entitlement_access_packages` | R | params |
| `microsoft_list_group_conversations` | R | group_id, params |
| `microsoft_list_group_members` | R | group_id, params |
| `microsoft_list_group_owners` | R | group_id, params |
| `microsoft_list_groups` | R | params |
| `microsoft_list_lifecycle_workflows` | R | params |
| `microsoft_list_organization` | R | params |
| `microsoft_list_permission_grant_policies` | R | params |
| `microsoft_list_plan_tasks` | R | params, plannerPlan_id |
| `microsoft_list_provisioning_logs` | R | params |
| `microsoft_list_role_assignments` | R | params |
| `microsoft_list_role_definitions` | R | params |
| `microsoft_list_service_principals` | R | params |
| `microsoft_list_token_issuance_policies` | R | params |
| `microsoft_list_token_lifetime_policies` | R | params |
| `microsoft_list_users` | R | params |
| `microsoft_remove_application_password` | D | app_id, confirm, data, params |
| `microsoft_remove_group_member` | D | confirm, group_id, member_id, params |
| `microsoft_restore_deleted_item` | W | object_id, params |
| `microsoft_update_application` | W | app_id, data, params |
| `microsoft_update_conditional_access_policy` | W | data, params, policy_id |
| `microsoft_update_group` | W | data, group_id, params |
| `microsoft_update_org_branding` | W | data, org_id, params |
| `microsoft_update_organization` | W | data, org_id, params |
| `microsoft_update_service_principal` | W | data, params, sp_id |
| `microsoft_verify_domain` | W | domain_id, params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `identity-access-review`: `microsoft_get_access_review` -> `microsoft_list_access_reviews`
- `identity-agreement`: `microsoft_get_agreement` -> `microsoft_list_agreements` -> `microsoft_create_agreement` -> `microsoft_delete_agreement`
- `identity-application`: `microsoft_get_application` -> `microsoft_list_applications` -> `microsoft_create_application` -> `microsoft_update_application` -> `microsoft_delete_application`
- `identity-application-password`: `microsoft_add_application_password` -> `microsoft_remove_application_password`
- `identity-conditional-access`: `microsoft_get_conditional_access_policy` -> `microsoft_list_conditional_access_policies` -> `microsoft_create_conditional_access_policy` -> `microsoft_update_conditional_access_policy` -> `microsoft_delete_conditional_access_policy`
- `identity-delegated-admin`: `microsoft_get_delegated_admin_relationship` -> `microsoft_list_delegated_admin_relationships`
- `identity-deleted-item`: `microsoft_list_deleted_items` -> `microsoft_restore_deleted_item`
- `identity-device`: `microsoft_get_device` -> `microsoft_list_devices` -> `microsoft_delete_device`
- `identity-directory-object`: `microsoft_get_directory_object` -> `microsoft_list_directory_objects`
- `identity-directory-role`: `microsoft_get_directory_role` -> `microsoft_list_directory_role_templates` -> `microsoft_list_directory_roles`
- `identity-domain`: `microsoft_get_domain` -> `microsoft_list_domains` -> `microsoft_create_domain` -> `microsoft_verify_domain` -> `microsoft_delete_domain`
- `identity-group`: `microsoft_get_group` -> `microsoft_list_groups` -> `microsoft_create_group` -> `microsoft_update_group` -> `microsoft_delete_group`
- `identity-group-member`: `microsoft_list_group_members` -> `microsoft_add_group_member` -> `microsoft_remove_group_member`
- `identity-org-branding`: `microsoft_get_org_branding` -> `microsoft_update_org_branding`
- `identity-organization`: `microsoft_get_organization` -> `microsoft_list_organization` -> `microsoft_update_organization`
- `identity-role-assignment`: `microsoft_get_role_assignment` -> `microsoft_list_role_assignments` -> `microsoft_create_role_assignment`
- `identity-role-definition`: `microsoft_get_role_definition` -> `microsoft_list_role_definitions`
- `identity-service-principal`: `microsoft_get_service_principal` -> `microsoft_list_service_principals` -> `microsoft_create_service_principal` -> `microsoft_update_service_principal` -> `microsoft_delete_service_principal`
- `identity-read-records-1`: `microsoft_get_admin_consent_policy` -> `microsoft_get_authorization_policy` -> `microsoft_get_me` -> `microsoft_list_domain_service_configuration_records` -> `microsoft_list_entitlement_access_packages` -> `microsoft_list_group_conversations` -> `microsoft_list_group_owners` -> `microsoft_list_lifecycle_workflows` -> `microsoft_list_permission_grant_policies` -> `microsoft_list_plan_tasks` -> `microsoft_list_provisioning_logs` -> `microsoft_list_token_issuance_policies`
- `identity-read-records-2`: `microsoft_list_token_lifetime_policies` -> `microsoft_list_users`
- `identity-change-records`: `microsoft_create_invitation`
