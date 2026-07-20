from typing import Any

from microsoft_agent.api._graph_models import (
    graph_model_from_dict,
)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_calendar_event(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = graph_model_from_dict(data, Event)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_calendar_event(
        self, event_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = graph_model_from_dict(data, Event)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def delete_calendar_event(
        self, event_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete calendar event."""
        try:
            await self.client.me.events.by_event_id(event_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_specific_calendar_event(
        self, calendar_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create specific calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = graph_model_from_dict(data, Event)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            event = graph_model_from_dict(data, Event)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            request_body = graph_model_from_dict(data, FindMeetingTimesPostRequestBody)

            request_config = FindMeetingTimesRequestBuilder.FindMeetingTimesRequestBuilderPostRequestConfiguration(
                options=[ResponseHandlerOption(NativeResponseHandler())]
            )
            native_response = await self.client.me.find_meeting_times.post(
                request_body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_outlook_contact(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Outlook contact."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.contact import Contact

        try:
            contact = graph_model_from_dict(data, Contact)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_outlook_contact(
        self, contact_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update Outlook contact."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.contact import Contact

        try:
            contact = graph_model_from_dict(data, Contact)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def delete_outlook_contact(
        self, contact_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete Outlook contact."""
        try:
            await self.client.me.contacts.by_contact_id(contact_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_online_meetings(self, params: dict | None = None) -> dict[str, Any]:
        """List online meetings for the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.online_meetings.online_meetings_request_builder import (
            OnlineMeetingsRequestBuilder,
        )

        query_params = OnlineMeetingsRequestBuilder.OnlineMeetingsRequestBuilderGetQueryParameters()
        if params:
            if "$filter" in params:
                query_params.filter = params["$filter"]
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")
            if "$orderby" in params:
                query_params.orderby = params["$orderby"].split(",")
            if "$search" in params:
                query_params.search = params["$search"]
            if "$skip" in params:
                query_params.skip = int(params["$skip"])
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$count" in params:
                query_params.count = str(params["$count"]).casefold() == "true"

        try:
            request_config = OnlineMeetingsRequestBuilder.OnlineMeetingsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            if params and "Accept-Language" in params:
                request_config.headers.add("Accept-Language", params["Accept-Language"])
            native_response = await self.client.me.online_meetings.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_online_meeting(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a new online meeting."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.online_meeting import OnlineMeeting

        try:
            meeting = graph_model_from_dict(data, OnlineMeeting)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_online_meeting(
        self, meeting_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update an online meeting."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.online_meeting import OnlineMeeting

        try:
            meeting = graph_model_from_dict(data, OnlineMeeting)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}
