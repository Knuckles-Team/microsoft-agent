import os
import sys
from typing import Any

from microsoft_agent.auth import AuthManager

CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite",
    "Tasks.ReadWrite",
    "Contacts.ReadWrite",
    "Group.ReadWrite.All",
    "Directory.Read.All",
    "Sites.Read.All",
    "Chat.Read",
    "ChatMessage.Read.All",
    "ChannelMessage.Read.All",
    "ServiceHealth.Read.All",
    "ServiceMessage.Read.All",
    "Domain.ReadWrite.All",
    "Organization.ReadWrite.All",
    "OnlineMeetings.ReadWrite",
    "CallRecords.Read.All",
    "Presence.Read.All",
    "User.Invite.All",
    "SecurityEvents.ReadWrite.All",
    "SecurityIncident.ReadWrite.All",
    "ThreatHunting.Read.All",
    "AuditLog.Read.All",
    "Reports.Read.All",
    "Application.ReadWrite.All",
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "IdentityRiskEvent.Read.All",
    "IdentityRiskyUser.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory",
    "EntitlementManagement.Read.All",
    "AccessReview.Read.All",
    "LifecycleWorkflows.Read.All",
]

# Only create global auth_manager if not in test mode
auth_manager: AuthManager | None
if not os.environ.get("TESTING"):
    auth_manager = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)
else:
    auth_manager = None

from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiAdmin(MicrosoftGraphApiBase):
    async def list_service_health(self, params: dict | None = None) -> dict[str, Any]:
        """List service health overviews."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.health_overviews.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.admin.service_announcement.health_overviews.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service health: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_service_health(
        self, service_name: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get service health for a specific service."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.health_overviews.by_service_health_id(
                service_name
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.health_overviews.by_service_health_id(
                service_name
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service health: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_service_health_issues(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List service health issues."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.issues.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.issues.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service health issues: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_service_health_issue(
        self, issue_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific service health issue."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.issues.by_service_health_issue_id(
                issue_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.issues.by_service_health_issue_id(
                issue_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service health issue: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_organization(self, params: dict | None = None) -> dict[str, Any]:
        """List organization properties."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.organization.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.organization.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing organization: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_organization(
        self, org_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get organization by ID."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.organization.by_organization_id(
                org_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.organization.by_organization_id(
                org_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting organization: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_organization(
        self, org_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update organization properties."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.organization import Organization

        try:
            org = Organization()
            if "displayName" in data:
                org.display_name = data["displayName"]
            request_config = self.client.organization.by_organization_id(
                org_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.organization.by_organization_id(
                org_id
            ).patch(org, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating organization: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_domains(self, params: dict | None = None) -> dict[str, Any]:
        """List tenant domains."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.domains.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing domains: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_domain(
        self, domain_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get domain details."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.domains.by_domain_id(
                domain_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.by_domain_id(domain_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting domain: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_domain(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add a domain to the tenant."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.domain import Domain

        try:
            domain = Domain()
            domain.id = data.get("id")
            request_config = self.client.domains.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.post(
                domain, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating domain: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_domain(
        self, domain_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a domain."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.domains.by_domain_id(
                domain_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.by_domain_id(domain_id).delete(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting domain: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def verify_domain(
        self, domain_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Verify domain ownership."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.domains.by_domain_id(
                domain_id
            ).verify.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.by_domain_id(
                domain_id
            ).verify.post(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error verifying domain: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_domain_service_configuration_records(
        self, domain_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List domain service configuration DNS records."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.domains.by_domain_id(
                domain_id
            ).service_configuration_records.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.domains.by_domain_id(
                domain_id
            ).service_configuration_records.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(
                f"Error listing domain service configuration records: {e}",
                file=sys.stderr,
            )
            return {"error": str(e)}

    async def list_security_alerts(self, params: dict | None = None) -> dict[str, Any]:
        """List security alerts (v2)."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.security.alerts_v2.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.alerts_v2.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing security alerts: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_security_alert(
        self, alert_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific security alert."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.security.alerts_v2.by_alert_id(
                alert_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.alerts_v2.by_alert_id(
                alert_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting security alert: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_security_alert(
        self, alert_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a security alert (e.g. change status, assign)."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.security.alert import Alert

        try:
            alert = Alert()
            if "status" in data:
                alert.status = data["status"]
            if "assignedTo" in data:
                alert.assigned_to = data["assignedTo"]
            if "classification" in data:
                alert.classification = data["classification"]
            if "determination" in data:
                alert.determination = data["determination"]
            request_config = self.client.security.alerts_v2.by_alert_id(
                alert_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.alerts_v2.by_alert_id(
                alert_id
            ).patch(alert, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating security alert: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_security_incidents(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List security incidents."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.security.incidents.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.incidents.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing security incidents: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_security_incident(
        self, incident_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific security incident."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.security.incidents.by_incident_id(
                incident_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.incidents.by_incident_id(
                incident_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting security incident: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_security_incident(
        self, incident_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a security incident."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.security.incident import Incident

        try:
            incident = Incident()
            if "status" in data:
                incident.status = data["status"]
            if "assignedTo" in data:
                incident.assigned_to = data["assignedTo"]
            if "classification" in data:
                incident.classification = data["classification"]
            if "determination" in data:
                incident.determination = data["determination"]
            request_config = self.client.security.incidents.by_incident_id(
                incident_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.incidents.by_incident_id(
                incident_id
            ).patch(incident, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating security incident: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def run_hunting_query(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Run an advanced hunting query."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.security.microsoft_graph_security_run_hunting_query.run_hunting_query_post_request_body import (
            RunHuntingQueryPostRequestBody,
        )

        try:
            body = RunHuntingQueryPostRequestBody()
            body.query = data.get("query")
            request_config = self.client.security.microsoft_graph_security_run_hunting_query.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.microsoft_graph_security_run_hunting_query.post(
                body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error running hunting query: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_directory_audits(self, params: dict | None = None) -> dict[str, Any]:
        """List directory audit logs."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.audit_logs.directory_audits.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.audit_logs.directory_audits.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing directory audits: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_directory_audit(
        self, audit_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific directory audit entry."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.audit_logs.directory_audits.by_directory_audit_id(
                    audit_id
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.audit_logs.directory_audits.by_directory_audit_id(
                    audit_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting directory audit: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_conditional_access_policies(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List conditional access policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity.conditional_access.policies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity.conditional_access.policies.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing conditional access policies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_conditional_access_policy(
        self, policy_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific conditional access policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting conditional access policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_conditional_access_policy(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a conditional access policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.conditional_access_policy import (
            ConditionalAccessPolicy,
        )

        try:
            policy = ConditionalAccessPolicy()
            if "displayName" in data:
                policy.display_name = data["displayName"]
            if "state" in data:
                policy.state = data["state"]
            request_config = self.client.identity.conditional_access.policies.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity.conditional_access.policies.post(
                    policy, request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating conditional access policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_conditional_access_policy(
        self, policy_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a conditional access policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.conditional_access_policy import (
            ConditionalAccessPolicy,
        )

        try:
            policy = ConditionalAccessPolicy()
            if "displayName" in data:
                policy.display_name = data["displayName"]
            if "state" in data:
                policy.state = data["state"]
            request_config = self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).patch(policy, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating conditional access policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_conditional_access_policy(
        self, policy_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a conditional access policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity.conditional_access.policies.by_conditional_access_policy_id(
                policy_id
            ).delete(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting conditional access policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_access_reviews(self, params: dict | None = None) -> dict[str, Any]:
        """List access review definitions."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_governance.access_reviews.definitions.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity_governance.access_reviews.definitions.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing access reviews: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_access_review(
        self, review_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific access review definition."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_governance.access_reviews.definitions.by_access_review_schedule_definition_id(
                review_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_governance.access_reviews.definitions.by_access_review_schedule_definition_id(
                review_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting access review: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_entitlement_access_packages(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List entitlement management access packages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_governance.entitlement_management.access_packages.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_governance.entitlement_management.access_packages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing access packages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_lifecycle_workflows(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List lifecycle management workflows."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_governance.lifecycle_workflows.workflows.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity_governance.lifecycle_workflows.workflows.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing lifecycle workflows: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_directory_roles(self, params: dict | None = None) -> dict[str, Any]:
        """List directory roles."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.directory_roles.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.directory_roles.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing directory roles: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_directory_role(
        self, role_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific directory role."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.directory_roles.by_directory_role_id(
                role_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.directory_roles.by_directory_role_id(
                role_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting directory role: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_directory_role_templates(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List directory role templates."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.directory_role_templates.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.directory_role_templates.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing directory role templates: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_authorization_policy(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """Get the authorization policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.policies.authorization_policy.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.authorization_policy.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting authorization policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_admin_consent_policy(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """Get the admin consent request policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.policies.admin_consent_request_policy.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.policies.admin_consent_request_policy.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting admin consent policy: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_role_definitions(self, params: dict | None = None) -> dict[str, Any]:
        """List role definitions."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.role_management.directory.role_definitions.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.role_management.directory.role_definitions.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing role definitions: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_role_definition(
        self, role_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific role definition."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.role_management.directory.role_definitions.by_unified_role_definition_id(
                role_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.role_management.directory.role_definitions.by_unified_role_definition_id(
                role_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting role definition: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_role_assignments(self, params: dict | None = None) -> dict[str, Any]:
        """List role assignments."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.role_management.directory.role_assignments.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.role_management.directory.role_assignments.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing role assignments: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_role_assignment(
        self, assignment_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific role assignment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.role_management.directory.role_assignments.by_unified_role_assignment_id(
                assignment_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.role_management.directory.role_assignments.by_unified_role_assignment_id(
                assignment_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting role assignment: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_role_assignment(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a role assignment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.unified_role_assignment import (
            UnifiedRoleAssignment,
        )

        try:
            assignment = UnifiedRoleAssignment()
            assignment.role_definition_id = data.get("roleDefinitionId")
            assignment.principal_id = data.get("principalId")
            assignment.directory_scope_id = data.get("directoryScopeId", "/")
            request_config = self.client.role_management.directory.role_assignments.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.role_management.directory.role_assignments.post(
                    assignment, request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating role assignment: {e}", file=sys.stderr)
            return {"error": str(e)}
