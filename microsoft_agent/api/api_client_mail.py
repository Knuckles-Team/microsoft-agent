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


class MicrosoftGraphApiMail(MicrosoftGraphApiBase):
    async def list_mail_messages(self, params: dict | None = None) -> dict[str, Any]:
        """List mail messages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
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
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )

        try:
            native_response = await self.client.me.messages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_mail_folders(self, params: dict | None = None) -> dict[str, Any]:
        """List mail folders."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.mail_folders.mail_folders_request_builder import (
            MailFoldersRequestBuilder,
        )

        query_params = (
            MailFoldersRequestBuilder.MailFoldersRequestBuilderGetQueryParameters()
        )
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$filter" in params:
                query_params.filter = params["$filter"]

        request_config = (
            MailFoldersRequestBuilder.MailFoldersRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )
        try:
            native_response = await self.client.me.mail_folders.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing mail folders: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_mail_folder_messages(
        self, mailFolder_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List messages in a specific folder."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$filter" in params:
                query_params.filter = params["$filter"]
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$search" in params:
                query_params.search = params["$search"]

        request_config = (
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )
        try:
            native_response = await self.client.me.mail_folders.by_mail_folder_id(
                mailFolder_id
            ).messages.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing folder messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_mail_message(
        self, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        query_params = (
            MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters()
        )
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")

        request_config = (
            MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )
        try:
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def send_mail(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send mail."""
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        try:
            request_body = SendMailPostRequestBody()
            message = Message()

            msg_data = data.get("message", {})
            message.subject = msg_data.get("subject")

            body_data = msg_data.get("body", {})
            body = ItemBody()
            body.content = body_data.get("content")
            body.content_type = (
                BodyType.Html
                if body_data.get("contentType") == "HTML"
                else BodyType.Text
            )
            message.body = body

            to_recipients = []
            for recipient in msg_data.get("toRecipients", []):
                rec = Recipient()
                email = EmailAddress()
                email_data = recipient.get("emailAddress", {})
                email.address = email_data.get("address")
                email.name = email_data.get("name")
                rec.email_address = email
                to_recipients.append(rec)
            message.to_recipients = to_recipients

            request_body.message = message
            request_body.save_to_sent_items = data.get("saveToSentItems", True)

            await self.client.me.send_mail.post(request_body)
            return {"status": "success"}
        except Exception as e:
            print(f"Error sending mail: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_draft_email(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create draft email."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.recipient import Recipient

        try:
            message = Message()
            message.subject = data.get("subject")

            body_data = data.get("body", {})
            if body_data:
                body = ItemBody()
                body.content = body_data.get("content")
                body.content_type = (
                    BodyType.Html
                    if body_data.get("contentType") == "HTML"
                    else BodyType.Text
                )
                message.body = body

            to_recipients = []
            for recipient in data.get("toRecipients", []):
                rec = Recipient()
                email = EmailAddress()
                email_data = recipient.get("emailAddress", {})
                email.address = email_data.get("address")
                email.name = email_data.get("name")
                rec.email_address = email
                to_recipients.append(rec)
            message.to_recipients = to_recipients

            request_config = self.client.me.messages.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.messages.post(
                message, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating draft: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_mail_message(
        self, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a message."""
        try:
            await self.client.me.messages.by_message_id(message_id).delete()
            return {"status": "success"}
        except Exception as e:
            return {"error": str(e)}

    async def move_mail_message(
        self, message_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Move a message to a folder."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import (
            MovePostRequestBody,
        )
        from msgraph.generated.users.item.messages.item.move.move_request_builder import (
            MoveRequestBuilder,
        )

        try:
            request_body = MovePostRequestBody()
            request_body.destination_id = data.get("destinationId")

            request_config = (
                MoveRequestBuilder.MoveRequestBuilderPostRequestConfiguration(
                    options=[ResponseHandlerOption(NativeResponseHandler())]
                )
            )
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).move.post(request_body, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error moving message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_mail_message(
        self, message_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update a message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.message import Message
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        message = Message()
        message.subject = data.get("subject")
        message.is_read = data.get("isRead")

        request_config = MessageItemRequestBuilder.MessageItemRequestBuilderPatchRequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())]
        )

        try:
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).patch(message, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            return {"error": str(e)}

    async def add_mail_attachment(
        self, message_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Add attachment to message."""
        import base64

        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.file_attachment import FileAttachment
        from msgraph.generated.users.item.messages.item.attachments.attachments_request_builder import (
            AttachmentsRequestBuilder,
        )

        try:
            attachment = FileAttachment()
            attachment.name = data.get("name")
            attachment.content_type = data.get("contentType")

            content_bytes = data.get("contentBytes")
            if content_bytes:
                attachment.content_bytes = base64.b64decode(content_bytes)

            request_config = AttachmentsRequestBuilder.AttachmentsRequestBuilderPostRequestConfiguration(
                options=[ResponseHandlerOption(NativeResponseHandler())]
            )
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).attachments.post(attachment, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error adding attachment: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_mail_attachments(
        self, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List attachments."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.attachments.attachments_request_builder import (
            AttachmentsRequestBuilder,
        )

        try:
            query_params = (
                AttachmentsRequestBuilder.AttachmentsRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = AttachmentsRequestBuilder.AttachmentsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).attachments.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing attachments: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_mail_attachment(
        self, message_id: str, attachment_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get attachment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.attachments.item.attachment_item_request_builder import (
            AttachmentItemRequestBuilder,
        )

        try:
            query_params = AttachmentItemRequestBuilder.AttachmentItemRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = AttachmentItemRequestBuilder.AttachmentItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            native_response = (
                await self.client.me.messages.by_message_id(message_id)
                .attachments.by_attachment_id(attachment_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting attachment: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_mail_attachment(
        self, message_id: str, attachment_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete attachment."""
        try:
            await (
                self.client.me.messages.by_message_id(message_id)
                .attachments.by_attachment_id(attachment_id)
                .delete()
            )
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting attachment: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_shared_mailbox_messages(
        self, user_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List messages in a shared mailbox."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        try:
            query_params = (
                MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")
                if "$filter" in params:
                    query_params.filter = params["$filter"]
                if "$top" in params:
                    query_params.top = int(params["$top"])
                if "$search" in params:
                    query_params.search = params["$search"]

            request_config = (
                MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = await self.client.users.by_user_id(user_id).messages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing shared mailbox messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_shared_mailbox_folder_messages(
        self, user_id: str, mailFolder_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List messages in a shared mailbox folder."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
            MessagesRequestBuilder,
        )

        try:
            query_params = (
                MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")
                if "$filter" in params:
                    query_params.filter = params["$filter"]
                if "$top" in params:
                    query_params.top = int(params["$top"])
                if "$search" in params:
                    query_params.search = params["$search"]

            request_config = (
                MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = (
                await self.client.users.by_user_id(user_id)
                .mail_folders.by_mail_folder_id(mailFolder_id)
                .messages.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing shared mailbox folder messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_shared_mailbox_message(
        self, user_id: str, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a message from a shared mailbox."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        try:
            query_params = (
                MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            native_response = (
                await self.client.users.by_user_id(user_id)
                .messages.by_message_id(message_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting shared mailbox message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def send_shared_mailbox_mail(
        self, user_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send mail from a shared mailbox."""
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.email_address import EmailAddress
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        try:
            request_body = SendMailPostRequestBody()
            message = Message()

            msg_data = data.get("message", {})
            message.subject = msg_data.get("subject")

            body_data = msg_data.get("body", {})
            body = ItemBody()
            body.content = body_data.get("content")
            body.content_type = (
                BodyType.Html
                if body_data.get("contentType") == "HTML"
                else BodyType.Text
            )
            message.body = body

            to_recipients = []
            for recipient in msg_data.get("toRecipients", []):
                rec = Recipient()
                email = EmailAddress()
                email_data = recipient.get("emailAddress", {})
                email.address = email_data.get("address")
                email.name = email_data.get("name")
                rec.email_address = email
                to_recipients.append(rec)
            message.to_recipients = to_recipients

            request_body.message = message
            request_body.save_to_sent_items = data.get("saveToSentItems", True)

            await self.client.users.by_user_id(user_id).send_mail.post(request_body)

            return {"status": "success"}
        except Exception as e:
            print(f"Error sending shared mailbox mail: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_chat_messages(
        self, chat_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List chat messages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.chats.by_chat_id(
                chat_id
            ).messages.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.chats.by_chat_id(chat_id).messages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing chat messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_chat_message(
        self, chat_id: str, chatMessage_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get chat message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting chat message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def send_chat_message(
        self, chat_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send chat message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.chat_message import ChatMessage
        from msgraph.generated.models.item_body import ItemBody

        try:
            message = ChatMessage()
            body = ItemBody()
            body.content = data.get("body", {}).get("content")
            message.body = body

            request_config = self.client.chats.by_chat_id(
                chat_id
            ).messages.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.chats.by_chat_id(chat_id).messages.post(
                message, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error sending chat message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_channel_messages(
        self, team_id: str, channel_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List channel messages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing channel messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_channel_message(
        self,
        team_id: str,
        channel_id: str,
        chatMessage_id: str,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Get channel message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(chatMessage_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(chatMessage_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting channel message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def send_channel_message(
        self,
        team_id: str,
        channel_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Send channel message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.chat_message import ChatMessage
        from msgraph.generated.models.item_body import ItemBody

        try:
            message = ChatMessage()
            body = ItemBody()
            body.content = data.get("body", {}).get("content")
            message.body = body

            request_config = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.post(message, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error sending channel message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_chat_message_replies(
        self, chat_id: str, chatMessage_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List chat message replies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing chat message replies: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def reply_to_chat_message(
        self,
        chat_id: str,
        chatMessage_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Reply to a chat message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.chat_message import ChatMessage
        from msgraph.generated.models.item_body import ItemBody

        try:
            message = ChatMessage()
            body = ItemBody()
            body.content = data.get("body", {}).get("content")
            message.body = body

            request_config = (
                self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.chats.by_chat_id(chat_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies.post(message, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error replying to chat message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_service_update_messages(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List service update messages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.messages.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.messages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service update messages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_service_update_message(
        self, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get a specific service update message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.admin.service_announcement.messages.by_service_update_message_id(
                message_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.messages.by_service_update_message_id(
                message_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service update message: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_email_activity_report(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get email activity user detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.reports.get_email_activity_user_detail_with_period(
                    period
                ).to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.reports.get_email_activity_user_detail_with_period(
                    period
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting email activity report: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_mailbox_usage_report(
        self, period: str = "D7", params: dict | None = None
    ) -> dict[str, Any]:
        """Get mailbox usage detail report."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.reports.get_mailbox_usage_detail_with_period(
                period
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = (
                await self.client.reports.get_mailbox_usage_detail_with_period(
                    period
                ).get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting mailbox usage report: {e}", file=sys.stderr)
            return {"error": str(e)}
