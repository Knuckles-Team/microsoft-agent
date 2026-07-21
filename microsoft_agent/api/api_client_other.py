from typing import Any

from microsoft_agent.api._graph_models import (
    decode_graph_base64,
    graph_model_from_dict,
)
from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiOther(MicrosoftGraphApiBase):
    async def get_me(self, params: dict | None = None) -> dict[str, Any]:
        """Get the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.user_item_request_builder import (
            UserItemRequestBuilder,
        )

        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters()
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")

        request_config = (
            UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )
        try:
            native_response = await self.client.me.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            body = graph_model_from_dict(data, QueryPostRequestBody)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            branding = graph_model_from_dict(data, OrganizationalBranding)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_subscription(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a subscription for change notifications."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subscription import Subscription

        try:
            subscription = graph_model_from_dict(data, Subscription)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_subscription(
        self, subscription_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update/renew a subscription."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subscription import Subscription

        try:
            subscription = graph_model_from_dict(data, Subscription)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_invitation(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an invitation for a guest user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.invitation import Invitation

        try:
            invitation = graph_model_from_dict(
                {"inviteRedirectUrl": "https://myapps.microsoft.com", **data},
                Invitation,
            )
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_application(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an application registration."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.application import Application

        try:
            app = graph_model_from_dict(data, Application)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_application(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.application import Application

        try:
            app = graph_model_from_dict(data, Application)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def add_application_password(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add a password credential to an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.applications.item.add_password.add_password_post_request_body import (
            AddPasswordPostRequestBody,
        )

        try:
            body = graph_model_from_dict(
                {"passwordCredential": data}, AddPasswordPostRequestBody
            )
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def remove_application_password(
        self, app_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Remove a password credential from an application."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.applications.item.remove_password.remove_password_post_request_body import (
            RemovePasswordPostRequestBody,
        )

        try:
            body = graph_model_from_dict(data, RemovePasswordPostRequestBody)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_service_principal(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.service_principal import ServicePrincipal

        try:
            sp = graph_model_from_dict(data, ServicePrincipal)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_service_principal(
        self, sp_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a service principal."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.service_principal import ServicePrincipal

        try:
            sp = graph_model_from_dict(data, ServicePrincipal)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_agreement(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create an agreement."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.agreement import Agreement

        try:
            agreement = graph_model_from_dict(data, Agreement)
            files_data = data.get("files")
            if files_data is not None:
                if not isinstance(files_data, list):
                    raise ValueError("files must be a list")
                agreement_files = agreement.files or []
                if len(agreement_files) != len(files_data):
                    raise ValueError("Agreement files could not be parsed")
                for index, (agreement_file, source) in enumerate(
                    zip(agreement_files, files_data, strict=True)
                ):
                    if not isinstance(source, dict):
                        raise ValueError("Each agreement file must be an object")
                    file_data = source.get("fileData")
                    if isinstance(file_data, dict) and "data" in file_data:
                        if agreement_file.file_data is None:
                            raise ValueError("Agreement fileData could not be parsed")
                        agreement_file.file_data.data = decode_graph_base64(
                            file_data["data"], f"files[{index}].fileData.data"
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_place(
        self, place_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a place."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.room import Room

        try:
            room = graph_model_from_dict(data, Room)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_printers(self, params: dict | None = None) -> dict[str, Any]:
        """List printers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print.printers.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print.printers.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_printer(
        self, printer_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific printer."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print.printers.by_printer_id(
                printer_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print.printers.by_printer_id(
                printer_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_print_jobs(
        self, printer_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List print jobs for a printer."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print.printers.by_printer_id(
                printer_id
            ).jobs.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print.printers.by_printer_id(
                printer_id
            ).jobs.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_print_job(
        self, printer_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a print job."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.print_job import PrintJob

        try:
            job = graph_model_from_dict(data, PrintJob)
            request_config = self.client.print.printers.by_printer_id(
                printer_id
            ).jobs.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print.printers.by_printer_id(
                printer_id
            ).jobs.post(job, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_subject_rights_request(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a subject rights request."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subject_rights_request import SubjectRightsRequest

        try:
            srr = graph_model_from_dict(data, SubjectRightsRequest)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_booking_appointment(
        self, business_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a booking appointment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.booking_appointment import BookingAppointment

        try:
            appointment = graph_model_from_dict(data, BookingAppointment)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            conn = graph_model_from_dict(data, ExternalConnection)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_print_document_upload_session(
        self,
        printer_id: str,
        print_job_id: str,
        print_document_id: str,
        document_name: str,
        content_type: str,
        size: int,
    ) -> dict[str, Any]:
        """Create the documented preauthenticated Universal Print upload session."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.print_document_upload_properties import (
            PrintDocumentUploadProperties,
        )
        from msgraph.generated.print.printers.item.jobs.item.documents.item.create_upload_session.create_upload_session_post_request_body import (
            CreateUploadSessionPostRequestBody,
        )

        try:
            properties = PrintDocumentUploadProperties(
                document_name=document_name,
                content_type=content_type,
                size=size,
            )
            body = CreateUploadSessionPostRequestBody(properties=properties)
            create_session = (
                self.client.print.printers.by_printer_id(printer_id)
                .jobs.by_print_job_id(print_job_id)
                .documents.by_print_document_id(print_document_id)
                .create_upload_session
            )
            request_config = create_session.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await create_session.post(
                body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def start_print_job(
        self, printer_id: str, print_job_id: str
    ) -> dict[str, Any]:
        """Start an uploaded Universal Print job."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            start = (
                self.client.print.printers.by_printer_id(printer_id)
                .jobs.by_print_job_id(print_job_id)
                .start
            )
            request_config = start.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await start.post(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def submit_print_document(
        self,
        printer_id: str,
        submission: Any,
        *,
        upload_transport: Any | None = None,
    ) -> dict[str, Any]:
        """Create, upload, and start one print job without exposing its upload URL."""
        from microsoft_agent.power_platform import HttpxAsyncHttpTransport
        from microsoft_agent.universal_print import (
            PrintDocumentSubmission,
            UniversalPrintUploader,
            validate_print_upload_url,
        )

        try:
            request = (
                submission
                if isinstance(submission, PrintDocumentSubmission)
                else PrintDocumentSubmission.model_validate(submission)
            )
            content = request.content_bytes()
        except (TypeError, ValueError) as exc:
            return {"error": str(exc), "stage": "validation"}

        job = await self.create_print_job(
            printer_id, {"configuration": request.configuration}
        )
        if "error" in job:
            return {"error": job["error"], "stage": "create_job"}
        job_id = job.get("id")
        documents = job.get("documents")
        if (
            not isinstance(job_id, str)
            or not isinstance(documents, list)
            or len(documents) != 1
            or not isinstance(documents[0], dict)
            or not isinstance(documents[0].get("id"), str)
        ):
            return {
                "error": "Microsoft Graph returned invalid print job metadata",
                "stage": "create_job",
            }
        document_id = documents[0]["id"]

        session = await self.create_print_document_upload_session(
            printer_id,
            job_id,
            document_id,
            request.document_name,
            request.content_type,
            len(content),
        )
        if "error" in session:
            return {"error": session["error"], "stage": "create_upload_session"}
        owned_transport: HttpxAsyncHttpTransport | None = None
        try:
            upload_url = validate_print_upload_url(session.get("uploadUrl"))
            if upload_transport is None:
                owned_transport = HttpxAsyncHttpTransport(
                    service="microsoft_graph",
                    tls_profile=self.auth_manager.graph_tls_profile,
                    tls_profile_ref=self.auth_manager.graph_tls_profile_ref,
                )
            uploader = UniversalPrintUploader(upload_transport or owned_transport)
            uploaded = await uploader.upload(upload_url, content)
        except (TypeError, ValueError, RuntimeError) as exc:
            return {"error": str(exc), "stage": "upload_document"}
        finally:
            if owned_transport is not None:
                owned_transport.close()

        status = await self.start_print_job(printer_id, job_id)
        if "error" in status:
            return {"error": status["error"], "stage": "start_job"}
        return {
            "jobId": job_id,
            "documentId": document_id,
            "document": uploaded,
            "status": status,
        }
