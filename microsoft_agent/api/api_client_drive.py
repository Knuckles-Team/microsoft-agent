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


class MicrosoftGraphApiDrive(MicrosoftGraphApiBase):
    async def list_drives(self, params: dict | None = None) -> dict[str, Any]:
        """List drives."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.drives.drives_request_builder import (
            DrivesRequestBuilder,
        )

        try:
            query_params = DrivesRequestBuilder.DrivesRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                DrivesRequestBuilder.DrivesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.me.drives.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing drives: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_drive_root_item(
        self, drive_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get drive root item."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.drives.by_drive_id(
                drive_id
            ).root.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.drives.by_drive_id(drive_id).root.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting drive root item: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_root_folder(
        self, drive_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Alias for get_drive_root_item."""
        return await self.get_drive_root_item(drive_id, params)

    async def list_folder_files(
        self, drive_id: str, driveItem_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List folder files."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.children.children_request_builder import (
            ChildrenRequestBuilder,
        )

        try:
            query_params = (
                ChildrenRequestBuilder.ChildrenRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                ChildrenRequestBuilder.ChildrenRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .children.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing folder files: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def download_onedrive_file_content(
        self, drive_id: str, driveItem_id: str, params: dict | None = None
    ) -> Any:
        """Download file content."""
        try:
            response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .content.get()
            )

            import base64

            if isinstance(response, bytes):
                return {"content": base64.b64encode(response).decode("utf-8")}
            return {"error": "Unexpected response type"}
        except Exception as e:
            print(f"Error downloading file content: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_onedrive_file(
        self, drive_id: str, driveItem_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete file."""
        try:
            await (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .delete()
            )
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting file: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def upload_file_content(
        self,
        drive_id: str,
        driveItem_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Upload file content."""
        import base64

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            content_bytes = data.get("contentBytes")
            if not content_bytes:
                return {"error": "No contentBytes provided"}

            body = base64.b64decode(content_bytes)

            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .content.to_put_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .content.put(body, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error uploading file content: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_sites(self, params: dict | None = None) -> dict[str, Any]:
        """List SharePoint sites."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.sites_request_builder import SitesRequestBuilder

        try:
            query_params = SitesRequestBuilder.SitesRequestBuilderGetQueryParameters()
            if params:
                if "$search" in params:
                    query_params.search = params["$search"]

            request_config = (
                SitesRequestBuilder.SitesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.sites.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing sites: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_site(
        self, site_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get SharePoint site."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.sites.by_site_id(
                site_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.sites.by_site_id(site_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting site: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_site_drives(
        self, site_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List drives for a SharePoint site."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.item.drives.drives_request_builder import (
            DrivesRequestBuilder,
        )

        try:
            query_params = DrivesRequestBuilder.DrivesRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                DrivesRequestBuilder.DrivesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.sites.by_site_id(site_id).drives.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing site drives: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_site_lists(
        self, site_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List lists for a SharePoint site."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.item.lists.lists_request_builder import (
            ListsRequestBuilder,
        )

        try:
            query_params = ListsRequestBuilder.ListsRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                ListsRequestBuilder.ListsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.sites.by_site_id(site_id).lists.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing site lists: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_site_list(
        self, site_id: str, list_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a SharePoint site list."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting site list: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sharepoint_site_by_path(
        self, site_id: str, path: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get SharePoint site by path."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.sites.by_site_id(site_id)
                .get_by_path(path)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.sites.by_site_id(site_id)
                .get_by_path(path)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting sharepoint site by path: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sharepoint_sites_delta(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """Get SharePoint sites delta."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.sites.delta.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.sites.delta.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting sharepoint sites delta: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_sharepoint_site_list_items(
        self, site_id: str, list_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List items in a SharePoint site list."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing site list items: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sharepoint_site_list_item(
        self,
        site_id: str,
        list_id: str,
        listItem_id: str,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Get an item in a SharePoint site list."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.by_list_item_id(listItem_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.sites.by_site_id(site_id)
                .lists.by_list_id(list_id)
                .items.by_list_item_id(listItem_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting site list item: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_group_drives(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List group drives."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).drives.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(group_id).drives.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing group drives: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_admin_sharepoint(self, params: dict | None = None) -> dict[str, Any]:
        """Get SharePoint admin settings."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.sharepoint.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.sharepoint.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting admin sharepoint: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_admin_sharepoint(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update SharePoint admin settings."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.sharepoint import Sharepoint

        try:
            sp = Sharepoint()
            request_config = (
                self.client.admin.sharepoint.to_patch_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.sharepoint.patch(
                sp, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating admin sharepoint: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_sharepoint_activity_report(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get SharePoint activity user detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.reports.get_share_point_activity_user_detail_with_period(
                    period
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.reports.get_share_point_activity_user_detail_with_period(
                period
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting SharePoint activity report: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_onedrive_usage_report(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get OneDrive usage account detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.reports.get_one_drive_usage_account_detail_with_period(
                    period
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.reports.get_one_drive_usage_account_detail_with_period(
                period
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting OneDrive usage report: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_permission_grant_policies(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List permission grant policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.policies.permission_grant_policies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.permission_grant_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing permission grant policies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_print_shares(self, params: dict | None = None) -> dict[str, Any]:
        """List print shares."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print_.shares.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print_.shares.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing print shares: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_file_storage_containers(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List file storage containers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.storage.file_storage.containers.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.storage.file_storage.containers.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing file storage containers: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_file_storage_container(
        self, container_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific file storage container."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.storage.file_storage.containers.by_file_storage_container_id(
                container_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.storage.file_storage.containers.by_file_storage_container_id(
                container_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting file storage container: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_file_storage_container(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a file storage container."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.file_storage_container import FileStorageContainer

        try:
            container = FileStorageContainer()
            container.display_name = data.get("displayName")
            container.container_type_id = data.get("containerTypeId")
            request_config = self.client.storage.file_storage.containers.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.storage.file_storage.containers.post(
                container, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating file storage container: {e}", file=sys.stderr)
            return {"error": str(e)}
