from typing import Any

from microsoft_agent.api._graph_models import (
    comma_separated_values,
    graph_model_from_dict,
    validated_sharepoint_delta_url,
)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_sharepoint_site_by_path(
        self, site_id: str, path: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get SharePoint site by path."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.sites.by_site_id(site_id)
                .get_by_path_with_path(path)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.sites.by_site_id(site_id)
                .get_by_path_with_path(path)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_sharepoint_sites_delta(
        self,
        params: dict | None = None,
        continuation_url: str | None = None,
        fetch_all: bool = False,
        max_pages: int = 100,
    ) -> dict[str, Any]:
        """Get or exhaust a SharePoint sites delta enumeration safely."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.delta.delta_request_builder import (
            DeltaRequestBuilder,
        )

        if not 1 <= max_pages <= 1000:
            return {"error": "max_pages must be between 1 and 1000"}

        try:
            supplied = dict(params or {})
            token = supplied.pop("token", None)
            if continuation_url and supplied:
                raise ValueError(
                    "Query parameters cannot be combined with a continuation URL"
                )
            if continuation_url and token is not None:
                raise ValueError("token cannot be combined with a continuation URL")
            unsupported = set(supplied) - {"$select", "$expand", "$top"}
            if unsupported:
                raise ValueError(
                    "Unsupported sites delta parameters: "
                    + ", ".join(sorted(unsupported))
                )

            if token is not None:
                if not isinstance(token, str) or not token or len(token) > 8192:
                    raise ValueError("sites delta token must be a non-empty string")
                from urllib.parse import urlencode

                query_values: list[tuple[str, str]] = [("token", token)]
                for name in ("$select", "$expand"):
                    if name in supplied:
                        value = supplied[name]
                        if not isinstance(value, str) or not value:
                            raise ValueError(f"{name} must be a non-empty string")
                        query_values.append((name, value))
                if "$top" in supplied:
                    top = int(supplied["$top"])
                    if not 1 <= top <= 999:
                        raise ValueError("$top must be between 1 and 999")
                    query_values.append(("$top", str(top)))
                continuation_url = (
                    "https://graph.microsoft.com/v1.0/sites/delta?"
                    + urlencode(query_values)
                )

            builder = self.client.sites.delta
            if continuation_url:
                builder = builder.with_url(
                    validated_sharepoint_delta_url(continuation_url)
                )

            query_params = None
            if not continuation_url:
                query_params = (
                    DeltaRequestBuilder.DeltaRequestBuilderGetQueryParameters()
                )
                if "$select" in supplied:
                    query_params.select = comma_separated_values(
                        supplied["$select"], "$select"
                    )
                if "$expand" in supplied:
                    query_params.expand = comma_separated_values(
                        supplied["$expand"], "$expand"
                    )
                if "$top" in supplied:
                    query_params.top = int(supplied["$top"])
                    if not 1 <= query_params.top <= 999:
                        raise ValueError("$top must be between 1 and 999")
        except (TypeError, ValueError) as exc:
            return {"error": str(exc)}

        try:
            request_config = (
                DeltaRequestBuilder.DeltaRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await builder.get(request_configuration=request_config)
            native_response.raise_for_status()
            payload = native_response.json()
            if not fetch_all:
                return payload

            values = payload.get("value")
            if not isinstance(values, list):
                return {"error": "Microsoft Graph returned an invalid delta page"}
            combined = list(values)
            pages_fetched = 1
            next_link = payload.get("@odata.nextLink")
            delta_link = payload.get("@odata.deltaLink")
            while next_link and pages_fetched < max_pages:
                next_builder = self.client.sites.delta.with_url(
                    validated_sharepoint_delta_url(next_link)
                )
                next_config = next_builder.to_get_request_configuration()
                next_config.options.append(
                    ResponseHandlerOption(NativeResponseHandler())
                )
                native_response = await next_builder.get(
                    request_configuration=next_config
                )
                native_response.raise_for_status()
                page = native_response.json()
                page_values = page.get("value")
                if not isinstance(page_values, list):
                    return {"error": "Microsoft Graph returned an invalid delta page"}
                combined.extend(page_values)
                pages_fetched += 1
                next_link = page.get("@odata.nextLink")
                delta_link = page.get("@odata.deltaLink")

            result: dict[str, Any] = {
                "value": combined,
                "pagesFetched": pages_fetched,
            }
            if delta_link:
                result["@odata.deltaLink"] = delta_link
            if next_link:
                result["@odata.nextLink"] = next_link
                result["partial"] = True
            return result
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_admin_sharepoint(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update SharePoint admin settings."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.sharepoint import Sharepoint

        try:
            sp = graph_model_from_dict(data, Sharepoint)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_print_shares(self, params: dict | None = None) -> dict[str, Any]:
        """List print shares."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.print.shares.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.print.shares.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_file_storage_container(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a file storage container."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.file_storage_container import FileStorageContainer

        try:
            container = graph_model_from_dict(data, FileStorageContainer)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_site_drive_by_id(
        self, site_id: str, drive_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a document library from a SharePoint site by drive ID."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.item.drives.item.drive_item_request_builder import (
            DriveItemRequestBuilder,
        )

        query_params = (
            DriveItemRequestBuilder.DriveItemRequestBuilderGetQueryParameters()
        )
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")

        request_config = (
            DriveItemRequestBuilder.DriveItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )

        try:
            native_response = (
                await self.client.sites.by_site_id(site_id)
                .drives.by_drive_id(drive_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_site_item(
        self, site_id: str, baseItem_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a base item addressed through a SharePoint site."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.item.items.item.base_item_item_request_builder import (
            BaseItemItemRequestBuilder,
        )

        query_params = (
            BaseItemItemRequestBuilder.BaseItemItemRequestBuilderGetQueryParameters()
        )
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")

        request_config = BaseItemItemRequestBuilder.BaseItemItemRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
            options=[ResponseHandlerOption(NativeResponseHandler())],
        )

        try:
            native_response = (
                await self.client.sites.by_site_id(site_id)
                .items.by_base_item_id(baseItem_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_site_items(
        self, site_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List base items addressed through a SharePoint site."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.sites.item.items.items_request_builder import (
            ItemsRequestBuilder,
        )

        query_params = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters()
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")
            if "$filter" in params:
                query_params.filter = params["$filter"]
            if "$orderby" in params:
                query_params.orderby = params["$orderby"].split(",")
            if "$search" in params:
                query_params.search = params["$search"]
            if "$skip" in params:
                query_params.skip = int(params["$skip"])
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$count" in params:
                query_params.count = str(params["$count"]).lower() == "true"

        request_config = ItemsRequestBuilder.ItemsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
            options=[ResponseHandlerOption(NativeResponseHandler())],
        )

        try:
            native_response = await self.client.sites.by_site_id(site_id).items.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}
