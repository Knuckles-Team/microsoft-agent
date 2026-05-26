import os
import sys
from typing import Any

from msgraph.generated.users.users_request_builder import UsersRequestBuilder

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


class MicrosoftGraphApiOther(MicrosoftGraphApiBase):
    async def get_me(self) -> dict[str, Any]:
        """Get the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())]
        )
        try:
            native_response = await self.client.me.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting me: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def search_query(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Search query."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.search.query.query_post_request_body import (
            QueryPostRequestBody,
        )

        try:
            body = QueryPostRequestBody()

            request_config = self.client.search.query.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.search.query.post(
                body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error performing search query: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_org_branding(
        self, org_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get organization branding."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.organization.by_organization_id(
                org_id
            ).branding.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.organization.by_organization_id(
                org_id
            ).branding.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting org branding: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_org_branding(
        self, org_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update organization branding."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.organizational_branding import (
            OrganizationalBranding,
        )

        try:
            branding = OrganizationalBranding()
            if "signInPageText" in data:
                branding.sign_in_page_text = data["signInPageText"]
            request_config = self.client.organization.by_organization_id(
                org_id
            ).branding.to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.organization.by_organization_id(
                org_id
            ).branding.patch(branding, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating org branding: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_subscriptions(self, params: dict | None = None) -> dict[str, Any]:
        """List active webhook subscriptions."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.subscriptions.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.subscriptions.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing subscriptions: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_subscription(
        self, subscription_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific subscription."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.subscriptions.by_subscription_id(
                subscription_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.subscriptions.by_subscription_id(
                subscription_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting subscription: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_subscription(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a subscription for change notifications."""
        import datetime

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subscription import Subscription

        try:
            subscription = Subscription()
            subscription.change_type = data.get("changeType")
            subscription.notification_url = data.get("notificationUrl")
            subscription.resource = data.get("resource")
            expiration = data.get("expirationDateTime")
            if expiration:
                subscription.expiration_date_time = datetime.datetime.fromisoformat(
                    expiration
                )
            subscription.client_state = data.get("clientState")
            request_config = self.client.subscriptions.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.subscriptions.post(
                subscription, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating subscription: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_subscription(
        self, subscription_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update/renew a subscription."""
        import datetime

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subscription import Subscription

        try:
            subscription = Subscription()
            expiration = data.get("expirationDateTime")
            if expiration:
                subscription.expiration_date_time = datetime.datetime.fromisoformat(
                    expiration
                )
            request_config = self.client.subscriptions.by_subscription_id(
                subscription_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.subscriptions.by_subscription_id(
                subscription_id
            ).patch(subscription, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating subscription: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_subscription(
        self, subscription_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a subscription."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.subscriptions.by_subscription_id(
                subscription_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.subscriptions.by_subscription_id(
                subscription_id
            ).delete(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting subscription: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_call_records(self, params: dict | None = None) -> dict[str, Any]:
        """List call records."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.communications.call_records.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.communications.call_records.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing call records: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_call_record(
        self, call_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific call record."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.communications.call_records.by_call_record_id(
                call_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.communications.call_records.by_call_record_id(
                    call_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting call record: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_invitation(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an invitation for a guest user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.invitation import Invitation

        try:
            invitation = Invitation()
            invitation.invited_user_email_address = data.get("invitedUserEmailAddress")
            invitation.invite_redirect_url = data.get(
                "inviteRedirectUrl", "https://myapps.microsoft.com"
            )
            if "invitedUserDisplayName" in data:
                invitation.invited_user_display_name = data["invitedUserDisplayName"]
            if "sendInvitationMessage" in data:
                invitation.send_invitation_message = data["sendInvitationMessage"]
            request_config = self.client.invitations.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.invitations.post(
                invitation, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating invitation: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_secure_scores(self, params: dict | None = None) -> dict[str, Any]:
        """List secure scores."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.security.secure_scores.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.secure_scores.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing secure scores: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_threat_intelligence_hosts(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List threat intelligence hosts."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.security.threat_intelligence.hosts.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.threat_intelligence.hosts.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing threat intelligence hosts: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_threat_intelligence_host(
        self, host_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific threat intelligence host."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.security.threat_intelligence.hosts.by_host_id(
                host_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.security.threat_intelligence.hosts.by_host_id(
                    host_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting threat intelligence host: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_sign_in_logs(self, params: dict | None = None) -> dict[str, Any]:
        """List sign-in logs."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.audit_logs.sign_ins.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.audit_logs.sign_ins.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing sign-in logs: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sign_in_log(
        self, sign_in_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific sign-in log entry."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.audit_logs.sign_ins.by_sign_in_id(
                sign_in_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.audit_logs.sign_ins.by_sign_in_id(
                sign_in_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting sign-in log: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_provisioning_logs(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List provisioning logs."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.audit_logs.provisioning.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.audit_logs.provisioning.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing provisioning logs: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_applications(self, params: dict | None = None) -> dict[str, Any]:
        """List app registrations."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.applications.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing applications: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_application(
        self, app_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.applications.by_application_id(
                app_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.by_application_id(
                app_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting application: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_application(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an application registration."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.application import Application

        try:
            app = Application()
            if "displayName" in data:
                app.display_name = data["displayName"]
            if "signInAudience" in data:
                app.sign_in_audience = data["signInAudience"]
            request_config = self.client.applications.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.post(
                app, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating application: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_application(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.application import Application

        try:
            app = Application()
            if "displayName" in data:
                app.display_name = data["displayName"]
            request_config = self.client.applications.by_application_id(
                app_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.by_application_id(
                app_id
            ).patch(app, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating application: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_application(
        self, app_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.applications.by_application_id(
                app_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.by_application_id(
                app_id
            ).delete(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting application: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def add_application_password(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add a password credential to an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.applications.item.add_password.add_password_post_request_body import (
            AddPasswordPostRequestBody,
        )
        from msgraph.generated.models.password_credential import PasswordCredential

        try:
            body = AddPasswordPostRequestBody()
            cred = PasswordCredential()
            if "displayName" in data:
                cred.display_name = data["displayName"]
            body.password_credential = cred
            request_config = self.client.applications.by_application_id(
                app_id
            ).add_password.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.by_application_id(
                app_id
            ).add_password.post(body, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error adding application password: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def remove_application_password(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Remove a password credential from an application."""
        import uuid

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.applications.item.remove_password.remove_password_post_request_body import (
            RemovePasswordPostRequestBody,
        )

        try:
            body = RemovePasswordPostRequestBody()
            body.key_id = uuid.UUID(data.get("keyId"))
            request_config = self.client.applications.by_application_id(
                app_id
            ).remove_password.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.applications.by_application_id(
                app_id
            ).remove_password.post(body, request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "password removed"}
        except Exception as e:
            print(f"Error removing application password: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_service_principals(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List service principals."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.service_principals.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.service_principals.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service principals: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_service_principal(
        self, sp_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.service_principals.by_service_principal_id(
                sp_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.service_principals.by_service_principal_id(sp_id).get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service principal: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_service_principal(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.service_principal import ServicePrincipal

        try:
            sp = ServicePrincipal()
            sp.app_id = data.get("appId")
            if "displayName" in data:
                sp.display_name = data["displayName"]
            request_config = (
                self.client.service_principals.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.service_principals.post(
                sp, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating service principal: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_service_principal(
        self, sp_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.service_principal import ServicePrincipal

        try:
            sp = ServicePrincipal()
            if "displayName" in data:
                sp.display_name = data["displayName"]
            request_config = self.client.service_principals.by_service_principal_id(
                sp_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.service_principals.by_service_principal_id(
                    sp_id
                ).patch(sp, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating service principal: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_service_principal(
        self, sp_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.service_principals.by_service_principal_id(
                sp_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.service_principals.by_service_principal_id(
                    sp_id
                ).delete(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting service principal: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_risk_detections(self, params: dict | None = None) -> dict[str, Any]:
        """List risk detections."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_protection.risk_detections.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_protection.risk_detections.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing risk detections: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_risk_detection(
        self, risk_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific risk detection."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_protection.risk_detections.by_risk_detection_id(
                    risk_id
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_protection.risk_detections.by_risk_detection_id(
                risk_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting risk detection: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_directory_objects(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List directory objects."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.directory_objects.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.directory_objects.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing directory objects: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_directory_object(
        self, object_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific directory object."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.directory_objects.by_directory_object_id(
                object_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.directory_objects.by_directory_object_id(
                    object_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting directory object: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_deleted_items(self, params: dict | None = None) -> dict[str, Any]:
        """List deleted directory items."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.directory.deleted_items.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.directory.deleted_items.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing deleted items: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def restore_deleted_item(
        self, object_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Restore a deleted directory item."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.directory.deleted_items.by_directory_object_id(
                object_id
            ).graph_user.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.directory.deleted_items.by_directory_object_id(
                    object_id
                ).restore.post(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error restoring deleted item: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_token_lifetime_policies(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List token lifetime policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.policies.token_lifetime_policies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.token_lifetime_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing token lifetime policies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_token_issuance_policies(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List token issuance policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.policies.token_issuance_policies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.token_issuance_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing token issuance policies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_devices(self, params: dict | None = None) -> dict[str, Any]:
        """List devices registered in the directory."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.devices.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.devices.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing devices: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_device(
        self, device_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific device."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.devices.by_device_id(
                device_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.devices.by_device_id(device_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting device: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_device(
        self, device_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a device."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.devices.by_device_id(
                device_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.devices.by_device_id(device_id).delete(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting device: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_managed_devices(self, params: dict | None = None) -> dict[str, Any]:
        """List managed devices."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.device_management.managed_devices.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.device_management.managed_devices.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing managed devices: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_managed_device(
        self, device_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific managed device."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.managed_devices.by_managed_device_id(
                    device_id
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.device_management.managed_devices.by_managed_device_id(
                device_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting managed device: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_device_compliance_policies(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List device compliance policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.device_management.device_compliance_policies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.device_management.device_compliance_policies.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing device compliance policies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_device_configurations(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List device configurations."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.device_management.device_configurations.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.device_management.device_configurations.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing device configurations: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def wipe_managed_device(
        self, device_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Wipe a managed device."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.managed_devices.by_managed_device_id(
                    device_id
                ).wipe.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.device_management.managed_devices.by_managed_device_id(
                device_id
            ).wipe.post(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "wipe initiated"}
        except Exception as e:
            print(f"Error wiping managed device: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def retire_managed_device(
        self, device_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Retire a managed device."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.managed_devices.by_managed_device_id(
                    device_id
                ).retire.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.device_management.managed_devices.by_managed_device_id(
                device_id
            ).retire.post(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "retire initiated"}
        except Exception as e:
            print(f"Error retiring managed device: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_education_classes(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List education classes."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.education.classes.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.education.classes.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing education classes: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_education_class(
        self, class_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific education class."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.education.classes.by_education_class_id(
                class_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.education.classes.by_education_class_id(
                class_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting education class: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_education_schools(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List education schools."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.education.schools.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.education.schools.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing education schools: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_education_school(
        self, school_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific education school."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.education.schools.by_education_school_id(
                school_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.education.schools.by_education_school_id(
                    school_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting education school: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_education_assignments(
        self, class_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List assignments for an education class."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.education.classes.by_education_class_id(
                class_id
            ).assignments.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.education.classes.by_education_class_id(
                class_id
            ).assignments.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing education assignments: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_agreements(self, params: dict | None = None) -> dict[str, Any]:
        """List agreements (terms of use)."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.agreements.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.agreements.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing agreements: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_agreement(
        self, agreement_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific agreement."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.agreements.by_agreement_id(
                agreement_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.agreements.by_agreement_id(
                agreement_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting agreement: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_agreement(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an agreement."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.agreement import Agreement

        try:
            agreement = Agreement()
            agreement.display_name = data.get("displayName")
            agreement.is_viewing_before_acceptance_required = data.get(
                "isViewingBeforeAcceptanceRequired", True
            )
            request_config = self.client.agreements.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.agreements.post(
                agreement, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating agreement: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_agreement(
        self, agreement_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete an agreement."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.agreements.by_agreement_id(
                agreement_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.agreements.by_agreement_id(
                agreement_id
            ).delete(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting agreement: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_rooms(self, params: dict | None = None) -> dict[str, Any]:
        """List rooms."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.places.graph_room.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.places.graph_room.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing rooms: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_room_lists(self, params: dict | None = None) -> dict[str, Any]:
        """List room lists."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.places.graph_room_list.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.places.graph_room_list.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing room lists: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_place(
        self, place_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific place."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.places.by_place_id(
                place_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.places.by_place_id(place_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting place: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_place(
        self, place_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a place."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.room import Room

        try:
            room = Room()
            if "displayName" in data:
                room.display_name = data["displayName"]
            if "capacity" in data:
                room.capacity = data["capacity"]
            request_config = self.client.places.by_place_id(
                place_id
            ).graph_room.to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.places.by_place_id(
                place_id
            ).graph_room.patch(room, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating place: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_printers(self, params: dict | None = None) -> dict[str, Any]:
        """List printers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print_.printers.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print_.printers.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing printers: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_printer(
        self, printer_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific printer."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print_.printers.by_printer_id(
                printer_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print_.printers.by_printer_id(
                printer_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting printer: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_print_jobs(
        self, printer_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List print jobs for a printer."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print_.printers.by_printer_id(
                printer_id
            ).jobs.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print_.printers.by_printer_id(
                printer_id
            ).jobs.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing print jobs: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_print_job(
        self, printer_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a print job."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.print_job import PrintJob

        try:
            job = PrintJob()
            request_config = self.client.print_.printers.by_printer_id(
                printer_id
            ).jobs.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print_.printers.by_printer_id(
                printer_id
            ).jobs.post(job, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating print job: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_subject_rights_requests(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List subject rights requests."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.privacy.subject_rights_requests.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.privacy.subject_rights_requests.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing subject rights requests: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_subject_rights_request(
        self, request_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific subject rights request."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.privacy.subject_rights_requests.by_subject_rights_request_id(
                request_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.privacy.subject_rights_requests.by_subject_rights_request_id(
                request_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting subject rights request: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_subject_rights_request(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a subject rights request."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subject_rights_request import SubjectRightsRequest

        try:
            srr = SubjectRightsRequest()
            srr.display_name = data.get("displayName")
            srr.description = data.get("description")
            request_config = self.client.privacy.subject_rights_requests.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.privacy.subject_rights_requests.post(
                srr, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating subject rights request: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_booking_businesses(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List booking businesses."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.solutions.booking_businesses.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.solutions.booking_businesses.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing booking businesses: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_booking_business(
        self, business_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific booking business."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting booking business: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_booking_appointments(
        self, business_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List booking appointments for a business."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).appointments.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).appointments.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing booking appointments: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_booking_appointment(
        self, business_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a booking appointment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.booking_appointment import BookingAppointment

        try:
            appointment = BookingAppointment()
            appointment.service_id = data.get("serviceId")
            appointment.customer_name = data.get("customerName")
            request_config = (
                self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).appointments.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.solutions.booking_businesses.by_booking_business_id(
                    business_id
                ).appointments.post(appointment, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating booking appointment: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_learning_providers(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List learning providers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.employee_experience.learning_providers.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.employee_experience.learning_providers.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing learning providers: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_learning_provider(
        self, provider_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific learning provider."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.employee_experience.learning_providers.by_learning_provider_id(
                provider_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.employee_experience.learning_providers.by_learning_provider_id(
                provider_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting learning provider: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_learning_course_activities(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List learning course activities for the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.employee_experience.learning_course_activities.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.me.employee_experience.learning_course_activities.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing learning course activities: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_external_connections(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List external connections."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.external.connections.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.external.connections.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing external connections: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_external_connection(
        self, connection_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific external connection."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.external.connections.by_external_connection_id(
                connection_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.external.connections.by_external_connection_id(
                    connection_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting external connection: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_external_connection(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an external connection."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.external_connectors.external_connection import (
            ExternalConnection,
        )

        try:
            conn = ExternalConnection()
            conn.id_ = data.get("id")
            conn.name = data.get("name")
            conn.description = data.get("description")
            request_config = (
                self.client.external.connections.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.external.connections.post(
                conn, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating external connection: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_external_connection(
        self, connection_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete an external connection."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.external.connections.by_external_connection_id(
                connection_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.external.connections.by_external_connection_id(
                    connection_id
                ).delete(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting external connection: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_sensitivity_labels(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List sensitivity labels."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.information_protection.policy.labels.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.information_protection.policy.labels.get(
                    request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing sensitivity labels: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sensitivity_label(
        self, label_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific sensitivity label."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.information_protection.policy.labels.by_information_protection_label_id(
                label_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.information_protection.policy.labels.by_information_protection_label_id(
                label_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting sensitivity label: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_delegated_admin_relationships(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List delegated admin relationships."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.tenant_relationships.delegated_admin_relationships.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.tenant_relationships.delegated_admin_relationships.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing delegated admin relationships: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_delegated_admin_relationship(
        self, rel_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific delegated admin relationship."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.tenant_relationships.delegated_admin_relationships.by_delegated_admin_relationship_id(
                rel_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.tenant_relationships.delegated_admin_relationships.by_delegated_admin_relationship_id(
                rel_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting delegated admin relationship: {e}", file=sys.stderr)
            return {"error": str(e)}
