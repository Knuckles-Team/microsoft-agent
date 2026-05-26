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
            print(f"Error listing users: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_current_user(self, params: dict | None = None) -> dict[str, Any]:
        """Get current user (alias for get_me)."""
        return await self.get_me()

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
            print(f"Error listing chats: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting chat: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing joined teams: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting team: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing team channels: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting team channel: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing team members: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing groups: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting group: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_group(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a new group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.group import Group

        try:
            group = Group()
            group.display_name = data.get("displayName")
            group.description = data.get("description")
            group.mail_enabled = data.get("mailEnabled", False)
            group.mail_nickname = data.get("mailNickname")
            group.security_enabled = data.get("securityEnabled", True)
            if "groupTypes" in data:
                group.group_types = data["groupTypes"]
            if "visibility" in data:
                group.visibility = data["visibility"]
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
            print(f"Error creating group: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_group(
        self, group_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.group import Group

        try:
            group = Group()
            if "displayName" in data:
                group.display_name = data["displayName"]
            if "description" in data:
                group.description = data["description"]
            if "visibility" in data:
                group.visibility = data["visibility"]
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
            print(f"Error updating group: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error deleting group: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing group members: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def add_group_member(
        self, group_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add a member to a group."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.reference_create import ReferenceCreate

        try:
            ref = ReferenceCreate()
            ref.odata_id = data.get(
                "@odata.id",
                f"https://graph.microsoft.com/v1.0/directoryObjects/{data.get('userId', data.get('id', ''))}",
            )
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
            print(f"Error adding group member: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error removing group member: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing group owners: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing group conversations: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing presences: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting presence: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting my presence: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting active users report: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting Teams user activity report: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing risky users: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error getting risky user: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error dismissing risky user: {e}", file=sys.stderr)
            return {"error": str(e)}

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
            print(f"Error listing education users: {e}", file=sys.stderr)
            return {"error": str(e)}
