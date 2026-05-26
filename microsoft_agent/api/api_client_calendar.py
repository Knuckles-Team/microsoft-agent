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


class MicrosoftGraphApiCalendar(MicrosoftGraphApiBase):
    async def list_calendar_events(
        self, params: dict | None = None, timezone: str | None = None
    ) -> dict[str, Any]:
        """List calendar events."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.events.events_request_builder import (
            EventsRequestBuilder,
        )

        try:
            query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")
                if "$filter" in params:
                    query_params.filter = params["$filter"]
                if "$top" in params:
                    query_params.top = int(params["$top"])

            request_config = (
                EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            if timezone:
                request_config.headers.add("Prefer", f'outlook.timezone="{timezone}"')

            native_response = await self.client.me.events.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing events: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_calendar_event(
        self,
        event_id: str,
        params: dict | None = None,
        timezone: str | None = None,
    ) -> dict[str, Any]:
        """Get calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.events.item.event_item_request_builder import (
            EventItemRequestBuilder,
        )

        try:
            query_params = (
                EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                EventItemRequestBuilder.EventItemRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            if timezone:
                request_config.headers.add("Prefer", f'outlook.timezone="{timezone}"')

            native_response = await self.client.me.events.by_event_id(event_id).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_calendar_event(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.item_body import ItemBody

        try:
            event = Event()
            event.subject = data.get("subject")

            body_data = data.get("body", {})
            if body_data:
                body = ItemBody()
                body.content = body_data.get("content")
                body.content_type = (
                    BodyType.Html
                    if body_data.get("contentType") == "HTML"
                    else BodyType.Text
                )
                event.body = body

            start_data = data.get("start", {})
            if start_data:
                start = DateTimeTimeZone()
                start.date_time = start_data.get("dateTime")
                start.time_zone = start_data.get("timeZone")
                event.start = start

            end_data = data.get("end", {})
            if end_data:
                end = DateTimeTimeZone()
                end.date_time = end_data.get("dateTime")
                end.time_zone = end_data.get("timeZone")
                event.end = end

            request_config = self.client.me.events.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.events.post(
                event, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_calendar_event(
        self, event_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = Event()
            event.subject = data.get("subject")

            request_config = self.client.me.events.by_event_id(
                event_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.events.by_event_id(event_id).patch(
                event, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_calendar_event(
        self, event_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete calendar event."""
        try:
            await self.client.me.events.by_event_id(event_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_calendars(self, params: dict | None = None) -> dict[str, Any]:
        """List calendars."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.calendars.calendars_request_builder import (
            CalendarsRequestBuilder,
        )

        try:
            query_params = (
                CalendarsRequestBuilder.CalendarsRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                CalendarsRequestBuilder.CalendarsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.me.calendars.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing calendars: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_calendar_view(
        self, params: dict | None = None, timezone: str | None = None
    ) -> dict[str, Any]:
        """Get calendar view."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
            CalendarViewRequestBuilder,
        )

        try:
            query_params = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters()
            if params:
                if "startDateTime" in params:
                    query_params.start_date_time = params["startDateTime"]
                if "endDateTime" in params:
                    query_params.end_date_time = params["endDateTime"]
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            if timezone:
                request_config.headers.add("Prefer", f'outlook.timezone="{timezone}"')

            native_response = await self.client.me.calendar_view.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting calendar view: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_specific_calendar_events(
        self,
        calendar_id: str,
        params: dict | None = None,
        timezone: str | None = None,
    ) -> dict[str, Any]:
        """List events for a specific calendar."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.calendars.item.events.events_request_builder import (
            EventsRequestBuilder,
        )

        try:
            query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")
                if "$filter" in params:
                    query_params.filter = params["$filter"]
                if "$top" in params:
                    query_params.top = int(params["$top"])

            request_config = (
                EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            if timezone:
                request_config.headers.add("Prefer", f'outlook.timezone="{timezone}"')

            native_response = await self.client.me.calendars.by_calendar_id(
                calendar_id
            ).events.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing specific calendar events: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_specific_calendar_event(
        self,
        calendar_id: str,
        event_id: str,
        params: dict | None = None,
        timezone: str | None = None,
    ) -> dict[str, Any]:
        """Get specific calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.calendars.item.events.item.event_item_request_builder import (
            EventItemRequestBuilder,
        )

        try:
            query_params = (
                EventItemRequestBuilder.EventItemRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                EventItemRequestBuilder.EventItemRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            if timezone:
                request_config.headers.add("Prefer", f'outlook.timezone="{timezone}"')

            native_response = (
                await self.client.me.calendars.by_calendar_id(calendar_id)
                .events.by_event_id(event_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting specific calendar event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_specific_calendar_event(
        self, calendar_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create specific calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = Event()
            event.subject = data.get("subject")

            request_config = self.client.me.calendars.by_calendar_id(
                calendar_id
            ).events.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.calendars.by_calendar_id(
                calendar_id
            ).events.post(event, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating specific calendar event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_specific_calendar_event(
        self,
        calendar_id: str,
        event_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Update specific calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = Event()
            event.subject = data.get("subject")

            request_config = (
                self.client.me.calendars.by_calendar_id(calendar_id)
                .events.by_event_id(event_id)
                .to_patch_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.calendars.by_calendar_id(calendar_id)
                .events.by_event_id(event_id)
                .patch(event, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating specific calendar event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_specific_calendar_event(
        self, calendar_id: str, event_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete specific calendar event."""
        try:
            await (
                self.client.me.calendars.by_calendar_id(calendar_id)
                .events.by_event_id(event_id)
                .delete()
            )
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting specific calendar event: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def find_meeting_times(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Find meeting times."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.find_meeting_times.find_meeting_times_post_request_body import (
            FindMeetingTimesPostRequestBody,
        )
        from msgraph.generated.users.item.find_meeting_times.find_meeting_times_request_builder import (
            FindMeetingTimesRequestBuilder,
        )

        try:
            request_body = FindMeetingTimesPostRequestBody()

            request_config = (
                FindMeetingTimesRequestBuilder.FindMeetingTimesPostRequestConfiguration(
                    options=[ResponseHandlerOption(NativeResponseHandler())]
                )
            )
            native_response = await self.client.me.find_meeting_times.post(
                request_body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error finding meeting times: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_outlook_contacts(self, params: dict | None = None) -> dict[str, Any]:
        """List Outlook contacts."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.contacts.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.contacts.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing outlook contacts: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_outlook_contact(
        self, contact_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Outlook contact."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.contacts.by_contact_id(
                contact_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.contacts.by_contact_id(
                contact_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting outlook contact: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_outlook_contact(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Outlook contact."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.contact import Contact
        from msgraph.generated.models.email_address import EmailAddress

        try:
            contact = Contact()
            contact.given_name = data.get("givenName")
            contact.surname = data.get("surname")

            emails = []
            for email_str in data.get("emailAddresses", []):
                email = EmailAddress()
                email.address = email_str
                emails.append(email)
            contact.email_addresses = emails

            request_config = self.client.me.contacts.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.contacts.post(
                contact, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating outlook contact: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_outlook_contact(
        self, contact_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update Outlook contact."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.contact import Contact

        try:
            contact = Contact()
            contact.given_name = data.get("givenName")
            contact.surname = data.get("surname")

            request_config = self.client.me.contacts.by_contact_id(
                contact_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.contacts.by_contact_id(
                contact_id
            ).patch(contact, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating outlook contact: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_outlook_contact(
        self, contact_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete Outlook contact."""
        try:
            await self.client.me.contacts.by_contact_id(contact_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting outlook contact: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_online_meetings(self, params: dict | None = None) -> dict[str, Any]:
        """List online meetings for the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.me.online_meetings.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.online_meetings.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing online meetings: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_online_meeting(
        self, meeting_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific online meeting."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting online meeting: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_online_meeting(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a new online meeting."""
        import datetime

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.online_meeting import OnlineMeeting

        try:
            meeting = OnlineMeeting()
            if "subject" in data:
                meeting.subject = data["subject"]
            if "startDateTime" in data:
                meeting.start_date_time = datetime.datetime.fromisoformat(
                    data["startDateTime"]
                )
            if "endDateTime" in data:
                meeting.end_date_time = datetime.datetime.fromisoformat(
                    data["endDateTime"]
                )
            request_config = (
                self.client.me.online_meetings.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.online_meetings.post(
                meeting, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating online meeting: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_online_meeting(
        self, meeting_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update an online meeting."""
        import datetime

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.online_meeting import OnlineMeeting

        try:
            meeting = OnlineMeeting()
            if "subject" in data:
                meeting.subject = data["subject"]
            if "startDateTime" in data:
                meeting.start_date_time = datetime.datetime.fromisoformat(
                    data["startDateTime"]
                )
            if "endDateTime" in data:
                meeting.end_date_time = datetime.datetime.fromisoformat(
                    data["endDateTime"]
                )
            request_config = self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).patch(meeting, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating online meeting: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_online_meeting(
        self, meeting_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete an online meeting."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).to_delete_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.me.online_meetings.by_online_meeting_id(
                meeting_id
            ).delete(request_configuration=request_config)
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting online meeting: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_virtual_events(self, params: dict | None = None) -> dict[str, Any]:
        """List virtual event townhalls."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.solutions.virtual_events.townhalls.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.solutions.virtual_events.townhalls.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing virtual events: {e}", file=sys.stderr)
            return {"error": str(e)}
