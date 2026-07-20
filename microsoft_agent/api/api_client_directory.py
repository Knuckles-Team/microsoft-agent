from typing import Any

from msgraph.generated.users.users_request_builder import UsersRequestBuilder

from microsoft_agent.api._graph_models import (
    graph_model_from_dict,
)
from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiDirectory(MicrosoftGraphApiBase):
    async def list_users(self, params: dict | None = None) -> dict[str, Any]:
        """List users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        query_params = UsersRequestBuilder.UsersRequestBuilderGetQueryParameters()

        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$filter" in params:
                query_params.filter = params["$filter"]
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$search" in params:
                query_params.search = params["$search"]
            if "$orderby" in params:
                query_params.orderby = params["$orderby"].split(",")
            if "$count" in params:
                query_params.count = params["$count"].lower() == "true"

        request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
            options=[ResponseHandlerOption(NativeResponseHandler())],
        )

        if params and "ConsistencyLevel" in params:
            request_config.headers.add("ConsistencyLevel", params["ConsistencyLevel"])

        try:
            native_response = await self.client.users.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_chats(self, params: dict | None = None) -> dict[str, Any]:
        """List user chats."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.chats.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.chats.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_chat(
        self, chat_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get chat."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.chats.by_chat_id(
                chat_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.chats.by_chat_id(chat_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_joined_teams(self, params: dict | None = None) -> dict[str, Any]:
        """List joined teams."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.joined_teams.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.joined_teams.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_team(
        self, team_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get team."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.teams.by_team_id(
                team_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.teams.by_team_id(team_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_team_channels(
        self, team_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List team channels."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.teams.by_team_id(
                team_id
            ).channels.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.teams.by_team_id(team_id).channels.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_team_channel(
        self, team_id: str, channel_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get team channel."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_team_members(
        self, team_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List team members."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.teams.by_team_id(
                team_id
            ).members.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.teams.by_team_id(team_id).members.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_groups(self, params: dict | None = None) -> dict[str, Any]:
        """List all Microsoft 365 groups and security groups."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder

        try:
            query_params = GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")
                if "$filter" in params:
                    query_params.filter = params["$filter"]
                if "$top" in params:
                    query_params.top = int(params["$top"])
                if "$search" in params:
                    query_params.search = params["$search"]
                if "$orderby" in params:
                    query_params.orderby = params["$orderby"].split(",")
                if "$count" in params:
                    query_params.count = params["$count"].lower() == "true"

            request_config = (
                GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            if params and "ConsistencyLevel" in params:
                request_config.headers.add(
                    "ConsistencyLevel", params["ConsistencyLevel"]
                )

            native_response = await self.client.groups.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_group(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(group_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_group(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a new group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.group import Group

        try:
            group = graph_model_from_dict(
                {"mailEnabled": False, "securityEnabled": True, **data}, Group
            )
            request_config = self.client.groups.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.post(
                group, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_group(
        self, group_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.group import Group

        try:
            group = graph_model_from_dict(data, Group)
            request_config = self.client.groups.by_group_id(
                group_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(group_id).patch(
                group, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def delete_group(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(group_id).delete(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_group_members(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List group members."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).members.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(
                group_id
            ).members.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def add_group_member(
        self, group_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add a member to a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.reference_create import ReferenceCreate

        try:
            reference_data = {
                "@odata.id": data.get(
                    "@odata.id",
                    "https://graph.microsoft.com/v1.0/directoryObjects/"
                    f"{data.get('userId', data.get('id', ''))}",
                )
            }
            ref = graph_model_from_dict(reference_data, ReferenceCreate)
            request_config = self.client.groups.by_group_id(
                group_id
            ).members.ref.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(
                group_id
            ).members.ref.post(ref, request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "member added"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def remove_group_member(
        self, group_id: str, member_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Remove a member from a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.groups.by_group_id(group_id)
                .members.by_directory_object_id(member_id)
                .ref.to_delete_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.groups.by_group_id(group_id)
                .members.by_directory_object_id(member_id)
                .ref.delete(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"status": "member removed"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_group_owners(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List group owners."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).owners.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(group_id).owners.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_group_conversations(
        self, group_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List group conversations."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.groups.by_group_id(
                group_id
            ).conversations.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.groups.by_group_id(
                group_id
            ).conversations.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_presences(self, params: dict | None = None) -> dict[str, Any]:
        """List presence information for users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.communications.presences.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.communications.presences.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_presence(
        self, user_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get presence for a specific user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.communications.presences.by_presence_id(
                user_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.communications.presences.by_presence_id(
                user_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_my_presence(self, params: dict | None = None) -> dict[str, Any]:
        """Get current user's presence."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.presence.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.presence.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_office365_active_users(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get Office 365 active user detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.reports.get_office365_active_user_detail_with_period(
                    period
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.reports.get_office365_active_user_detail_with_period(
                    period
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_teams_user_activity(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get Teams user activity detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.reports.get_teams_user_activity_user_detail_with_period(
                    period
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.reports.get_teams_user_activity_user_detail_with_period(
                period
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_risky_users(self, params: dict | None = None) -> dict[str, Any]:
        """List risky users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.identity_protection.risky_users.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_protection.risky_users.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_risky_user(
        self, user_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific risky user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_protection.risky_users.by_risky_user_id(
                    user_id
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity_protection.risky_users.by_risky_user_id(
                    user_id
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def dismiss_risky_user(
        self, user_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Dismiss a risky user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.identity_protection.risky_users.dismiss.dismiss_post_request_body import (
            DismissPostRequestBody,
        )

        try:
            body = DismissPostRequestBody()
            body.user_ids = [user_id]
            request_config = self.client.identity_protection.risky_users.dismiss.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.identity_protection.risky_users.dismiss.post(
                    body, request_configuration=request_config
                )
            )
            native_response.raise_for_status()
            return {"status": "dismissed"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_education_users(self, params: dict | None = None) -> dict[str, Any]:
        """List education users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.education.users.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.education.users.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}
