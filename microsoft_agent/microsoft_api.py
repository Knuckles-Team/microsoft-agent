import os
from typing import Dict, List, Optional, Any
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder

from microsoft_agent.auth import AuthManager
from microsoft_agent.credential_adapter import AuthManagerCredential

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

auth_manager = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)


class MicrosoftGraphApi:
    def __init__(self, auth_manager: AuthManager):
        self.auth_manager = auth_manager
        self.credential = AuthManagerCredential(auth_manager)
        self.client = GraphServiceClient(
            credentials=self.credential,
            # scopes=SCOPES # The SDK uses .default usually, but we can try passing scopes if needed.
            # actually GraphServiceClient constructor takes scopes.
            scopes=SCOPES,
        )

    def login(self, force: bool = False) -> str:
        """Authenticate with Microsoft."""
        if not force:
            token = self.auth_manager.get_token()
            if token:
                return "Already authenticated."

        def callback(msg):
            print(msg)
            # We could also use logging

        return self.auth_manager.acquire_token_by_device_code(callback=callback)

    def logout(self) -> str:
        """Logout."""
        self.auth_manager.logout()
        return "Logged out."

    def verify_login(self) -> str:
        """Verify login status."""
        token = self.auth_manager.get_token()
        if token:
            account = self.auth_manager.get_current_account()
            username = account.get("username") if account else "Unknown"
            return f"Authenticated as {username}"
        return "Not authenticated."

    def list_accounts(self) -> List[Dict[str, Any]]:
        """List accounts."""
        return self.auth_manager.list_accounts()

    def search_tools(self, query: str, limit: int = 10) -> List[str]:
        """Search methods in this class."""
        # Simple implementation searching method names
        matches = []
        for name in dir(self):
            if name.startswith("_"):
                continue
            if query.lower() in name.lower():
                matches.append(name)
            if len(matches) >= limit:
                break
        return matches

    async def list_mail_messages(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing messages: {e}")
            return {"error": str(e)}

    async def list_mail_folders(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing mail folders: {e}")
            return {"error": str(e)}

    async def list_mail_folder_messages(
        self, mailFolder_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing folder messages: {e}")
            return {"error": str(e)}

    async def get_mail_message(
        self, message_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting message: {e}")
            return {"error": str(e)}

    async def get_me(self) -> Dict[str, Any]:
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
            print(f"Error getting me: {e}")
            return {"error": str(e)}

    async def send_mail(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Send mail."""
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress

        # We receive a dict 'data' which should match the expected structure or we map it.
        # The SDK expects SendMailPostRequestBody which contains 'message' and 'saveToSentItems'.
        # If 'data' is just the message, we assume saveToSentItems=True or handle it.

        # NOTE: The previous API expected 'data' to be the body of the request.
        # But constructing the SDK objects is cleaner.
        # However, for MCP, it's easier if we accept the dict and try to deserialize it
        # OR just construct it from known fields.
        # Given "leverage the sdk", we should use the models if possible,
        # but manual mapping from dict is safer than trusting input structure matches SDK internals exactly.

        # Let's support a simplified input format compatible with previous tool usage if possible,
        # or just the raw Graph API JSON structure.
        # The best approach for MCP is usually accepting the Graph API JSON structure.
        # The user provided example uses SDK models.

        # I will implement a helper to convert dict to SendMailPostRequestBody if possible,
        # otherwise I will construct it manually from the dict.

        # Simplified:
        # data = { "message": { "subject": "...", "body": { "content": "..." }, "toRecipients": [...] } }

        # We can use the SDK's serialization to create the object from dict?
        # Only if we use the ParseNode factory... complex.
        # We'll map manually for common fields.

        try:
            # Basic mapping
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
            print(f"Error sending mail: {e}")
            return {"error": str(e)}

    async def create_draft_email(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create draft email."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress

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
            print(f"Error creating draft: {e}")
            return {"error": str(e)}

    async def list_users(self, params: Optional[Dict] = None) -> Dict[str, Any]:
        """List users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        # Parsing params for request configuration if needed,
        # but NativeResponseHandler returns raw response so we might just pass params?
        # The SDK's request builder expects QueryParameters object.

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
            print(f"Error listing users: {e}")
            return {"error": str(e)}

    async def delete_mail_message(
        self, message_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete a message."""
        try:
            await self.client.me.messages.by_message_id(message_id).delete()
            return {"status": "success"}
        except Exception as e:
            return {"error": str(e)}

    async def move_mail_message(
        self, message_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Move a message to a folder."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import (
            MovePostRequestBody,
        )
        from msgraph.generated.users.item.messages.item.move.move_post_request_builder import (
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
            print(f"Error moving message: {e}")
            return {"error": str(e)}

    async def update_mail_message(
        self, message_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Update a message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.message import Message
        from msgraph.generated.users.item.messages.item.message_item_request_builder import (
            MessageItemRequestBuilder,
        )

        # Construct message object from data
        message = Message()
        message.subject = data.get("subject")
        message.is_read = data.get("isRead")
        # Add other fields as needed

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
        self, message_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Add attachment to message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.file_attachment import FileAttachment
        from msgraph.generated.users.item.messages.item.attachments.attachments_request_builder import (
            AttachmentsRequestBuilder,
        )
        import base64

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
            print(f"Error adding attachment: {e}")
            return {"error": str(e)}

    async def list_mail_attachments(
        self, message_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing attachments: {e}")
            return {"error": str(e)}

    async def get_mail_attachment(
        self, message_id: str, attachment_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get attachment."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.messages.item.attachments.item.attachment_item_request_builder import (
            AttachmentItemRequestBuilder,
        )

        try:
            query_params = (
                AttachmentItemRequestBuilder.AttachmentItemRequestBuilderGetQueryParameters()
            )
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
            print(f"Error getting attachment: {e}")
            return {"error": str(e)}

    async def delete_mail_attachment(
        self, message_id: str, attachment_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete attachment."""
        try:
            await self.client.me.messages.by_message_id(
                message_id
            ).attachments.by_attachment_id(attachment_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting attachment: {e}")
            return {"error": str(e)}

    async def list_shared_mailbox_messages(
        self, user_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing shared mailbox messages: {e}")
            return {"error": str(e)}

    async def list_shared_mailbox_folder_messages(
        self, user_id: str, mailFolder_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing shared mailbox folder messages: {e}")
            return {"error": str(e)}

    async def get_shared_mailbox_message(
        self, user_id: str, message_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting shared mailbox message: {e}")
            return {"error": str(e)}

    async def send_shared_mailbox_mail(
        self, user_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Send mail from a shared mailbox."""
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )
        from msgraph.generated.models.message import Message
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress

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
            # send_mail returns None on success (202 Accepted)
            return {"status": "success"}
        except Exception as e:
            print(f"Error sending shared mailbox mail: {e}")
            return {"error": str(e)}

    async def list_calendar_events(
        self, params: Optional[Dict] = None, timezone: Optional[str] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing events: {e}")
            return {"error": str(e)}

    async def get_calendar_event(
        self,
        event_id: str,
        params: Optional[Dict] = None,
        timezone: Optional[str] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error getting event: {e}")
            return {"error": str(e)}

    async def create_calendar_event(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone

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
            print(f"Error creating event: {e}")
            return {"error": str(e)}

    async def update_calendar_event(
        self, event_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Update calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = Event()
            event.subject = data.get("subject")
            # ... map other fields as needed ...

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
            print(f"Error updating event: {e}")
            return {"error": str(e)}

    async def delete_calendar_event(
        self, event_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete calendar event."""
        try:
            await self.client.me.events.by_event_id(event_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting event: {e}")
            return {"error": str(e)}

    async def list_calendars(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing calendars: {e}")
            return {"error": str(e)}

    async def get_calendar_view(
        self, params: Optional[Dict] = None, timezone: Optional[str] = None
    ) -> Dict[str, Any]:
        """Get calendar view."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import (
            CalendarViewRequestBuilder,
        )

        try:
            query_params = (
                CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters()
            )
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
            print(f"Error getting calendar view: {e}")
            return {"error": str(e)}

    async def list_specific_calendar_events(
        self,
        calendar_id: str,
        params: Optional[Dict] = None,
        timezone: Optional[str] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error listing specific calendar events: {e}")
            return {"error": str(e)}

    async def get_specific_calendar_event(
        self,
        calendar_id: str,
        event_id: str,
        params: Optional[Dict] = None,
        timezone: Optional[str] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error getting specific calendar event: {e}")
            return {"error": str(e)}

    async def create_specific_calendar_event(
        self, calendar_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create specific calendar event."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.event import Event

        try:
            event = Event()
            event.subject = data.get("subject")
            # ... map other fields ...

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
            print(f"Error creating specific calendar event: {e}")
            return {"error": str(e)}

    async def update_specific_calendar_event(
        self,
        calendar_id: str,
        event_id: str,
        data: Dict[str, Any],
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error updating specific calendar event: {e}")
            return {"error": str(e)}

    async def delete_specific_calendar_event(
        self, calendar_id: str, event_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete specific calendar event."""
        try:
            await self.client.me.calendars.by_calendar_id(
                calendar_id
            ).events.by_event_id(event_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting specific calendar event: {e}")
            return {"error": str(e)}

    async def find_meeting_times(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            # ... map data to request_body ...
            # This is a complex body, for now we assume it's passed fairly rawly if possible,
            # or we map minimal fields if we want to be safe.

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
            print(f"Error finding meeting times: {e}")
            return {"error": str(e)}

    async def list_drives(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing drives: {e}")
            return {"error": str(e)}

    async def get_drive_root_item(
        self, drive_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting drive root item: {e}")
            return {"error": str(e)}

    async def get_root_folder(
        self, drive_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Alias for get_drive_root_item."""
        return await self.get_drive_root_item(drive_id, params)

    async def list_folder_files(
        self, drive_id: str, driveItem_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing folder files: {e}")
            return {"error": str(e)}

    async def download_onedrive_file_content(
        self, drive_id: str, driveItem_id: str, params: Optional[Dict] = None
    ) -> Any:
        """Download file content."""
        try:
            # content request returns bytes usually.
            response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(driveItem_id)
                .content.get()
            )
            # If we want raw response, we might need NativeResponseHandler.

            # FastMCP expects something serializable if it's a tool return.
            # Usually we return base64 for file content.
            import base64

            if isinstance(response, bytes):
                return {"content": base64.b64encode(response).decode("utf-8")}
            return {"error": "Unexpected response type"}
        except Exception as e:
            print(f"Error downloading file content: {e}")
            return {"error": str(e)}

    async def delete_onedrive_file(
        self, drive_id: str, driveItem_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete file."""
        try:
            await self.client.drives.by_drive_id(drive_id).items.by_drive_item_id(
                driveItem_id
            ).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting file: {e}")
            return {"error": str(e)}

    async def upload_file_content(
        self,
        drive_id: str,
        driveItem_id: str,
        data: Dict[str, Any],
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
        """Upload file content."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        import base64

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
            print(f"Error uploading file content: {e}")
            return {"error": str(e)}
            return {"error": str(e)}

    async def list_sites(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing sites: {e}")
            return {"error": str(e)}

    async def get_site(
        self, site_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting site: {e}")
            return {"error": str(e)}

    async def list_site_drives(
        self, site_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing site drives: {e}")
            return {"error": str(e)}

    async def list_site_lists(
        self, site_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing site lists: {e}")
            return {"error": str(e)}

    async def get_site_list(
        self, site_id: str, list_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting site list: {e}")
            return {"error": str(e)}

    async def get_excel_workbook(
        self, drive_id: str, item_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get Excel workbook."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel workbook: {e}")
            return {"error": str(e)}

    async def list_excel_worksheets(
        self, drive_id: str, item_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Excel worksheets."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.worksheets.worksheets_request_builder import (
            WorksheetsRequestBuilder,
        )

        try:
            query_params = (
                WorksheetsRequestBuilder.WorksheetsRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = WorksheetsRequestBuilder.WorksheetsRequestBuilderGetQueryParameters().WorksheetsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing excel worksheets: {e}")
            return {"error": str(e)}

    async def get_excel_worksheet(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
        """Get Excel worksheet."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel worksheet: {e}")
            return {"error": str(e)}

    async def list_excel_tables(
        self, drive_id: str, item_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Excel tables."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.tables.tables_request_builder import (
            TablesRequestBuilder,
        )

        try:
            query_params = TablesRequestBuilder.TablesRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                TablesRequestBuilder.TablesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing excel tables: {e}")
            return {"error": str(e)}

    async def get_excel_table(
        self, drive_id: str, item_id: str, table_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get Excel table."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.by_workbook_table_id(table_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.by_workbook_table_id(table_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel table: {e}")
            return {"error": str(e)}

    async def list_onenote_notebook_sections(
        self, notebook_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Onenote notebook sections."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.onenote.notebooks.by_notebook_id(
                notebook_id
            ).sections.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.onenote.notebooks.by_notebook_id(
                notebook_id
            ).sections.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing onenote sections: {e}")
            return {"error": str(e)}

    async def list_onenote_section_pages(
        self, onenoteSection_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Onenote section pages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.onenote.sections.by_onenote_section_id(
                onenoteSection_id
            ).pages.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.onenote.sections.by_onenote_section_id(
                    onenoteSection_id
                ).pages.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing onenote pages: {e}")
            return {"error": str(e)}

    async def get_onenote_page_content(
        self, onenotePage_id: str, params: Optional[Dict] = None
    ) -> Any:
        """Get Onenote page content."""
        try:
            response = await self.client.me.onenote.pages.by_onenote_page_id(
                onenotePage_id
            ).content.get()
            if isinstance(response, bytes):
                return {"content": response.decode("utf-8")}
            return {"error": "Unexpected response type"}
        except Exception as e:
            print(f"Error getting onenote page content: {e}")
            return {"error": str(e)}

    async def create_onenote_page(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create Onenote page."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            # Note: create_onenote_page usually takes multipart/form-data or HTML.
            # For now we use the HTML body if provided.
            html_content = data.get("content", "")

            request_config = (
                self.client.me.onenote.pages.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.onenote.pages.post(
                html_content.encode("utf-8"), request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating onenote page: {e}")
            return {"error": str(e)}

    async def list_todo_task_lists(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Todo task lists."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.todo.lists.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing todo task lists: {e}")
            return {"error": str(e)}

    async def list_todo_tasks(
        self, todoTaskList_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List Todo tasks."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing todo tasks: {e}")
            return {"error": str(e)}

    async def get_todo_task(
        self, todoTaskList_id: str, todoTask_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting todo task: {e}")
            return {"error": str(e)}

    async def create_todo_task(
        self, todoTaskList_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.todo_task import TodoTask

        try:
            task = TodoTask()
            task.title = data.get("title")
            # ... map other fields ...

            request_config = self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.post(task, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating todo task: {e}")
            return {"error": str(e)}

    async def update_todo_task(
        self,
        todoTaskList_id: str,
        todoTask_id: str,
        data: Dict[str, Any],
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
        """Update Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.todo_task import TodoTask

        try:
            task = TodoTask()
            task.title = data.get("title")

            request_config = (
                self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .to_patch_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .patch(task, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating todo task: {e}")
            return {"error": str(e)}

    async def delete_todo_task(
        self, todoTaskList_id: str, todoTask_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete Todo task."""
        try:
            await self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.by_todo_task_id(todoTask_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting todo task: {e}")
            return {"error": str(e)}

    async def list_planner_tasks(self, params: Optional[Dict] = None) -> Dict[str, Any]:
        """List Planner tasks."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.planner.tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.planner.tasks.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing planner tasks: {e}")
            return {"error": str(e)}

    async def get_planner_plan(
        self, plannerPlan_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get Planner plan."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting planner plan: {e}")
            return {"error": str(e)}

    async def list_plan_tasks(
        self, plannerPlan_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List tasks for a Planner plan."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).tasks.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing plan tasks: {e}")
            return {"error": str(e)}

    async def get_planner_task(
        self, plannerTask_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting planner task: {e}")
            return {"error": str(e)}

    async def create_planner_task(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            task = PlannerTask()
            task.title = data.get("title")
            task.plan_id = data.get("planId")
            # ... map other fields ...

            request_config = self.client.planner.tasks.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.post(
                task, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating planner task: {e}")
            return {"error": str(e)}

    async def update_planner_task(
        self, plannerTask_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Update Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            task = PlannerTask()
            task.title = data.get("title")

            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).patch(task, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating planner task: {e}")
            return {"error": str(e)}

    async def update_planner_task_details(
        self, plannerTask_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Update Planner task details."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task_details import PlannerTaskDetails

        try:
            details = PlannerTaskDetails()
            details.description = data.get("description")

            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.patch(details, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating planner task details: {e}")
            return {"error": str(e)}

    async def list_outlook_contacts(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing outlook contacts: {e}")
            return {"error": str(e)}

    async def get_outlook_contact(
        self, contact_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting outlook contact: {e}")
            return {"error": str(e)}

    async def create_outlook_contact(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating outlook contact: {e}")
            return {"error": str(e)}

    async def update_outlook_contact(
        self, contact_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating outlook contact: {e}")
            return {"error": str(e)}

    async def delete_outlook_contact(
        self, contact_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Delete Outlook contact."""
        try:
            await self.client.me.contacts.by_contact_id(contact_id).delete()
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting outlook contact: {e}")
            return {"error": str(e)}

    async def get_current_user(self, params: Optional[Dict] = None) -> Dict[str, Any]:
        """Get current user (alias for get_me)."""
        return await self.get_me()

    async def list_chats(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing chats: {e}")
            return {"error": str(e)}

    async def get_chat(
        self, chat_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting chat: {e}")
            return {"error": str(e)}

    async def list_chat_messages(
        self, chat_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing chat messages: {e}")
            return {"error": str(e)}

    async def get_chat_message(
        self, chat_id: str, chatMessage_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting chat message: {e}")
            return {"error": str(e)}

    async def send_chat_message(
        self, chat_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error sending chat message: {e}")
            return {"error": str(e)}

    async def list_joined_teams(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing joined teams: {e}")
            return {"error": str(e)}

    async def get_team(
        self, team_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting team: {e}")
            return {"error": str(e)}

    async def list_team_channels(
        self, team_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing team channels: {e}")
            return {"error": str(e)}

    async def get_team_channel(
        self, team_id: str, channel_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting team channel: {e}")
            return {"error": str(e)}

    async def list_channel_messages(
        self, team_id: str, channel_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing channel messages: {e}")
            return {"error": str(e)}

    async def get_channel_message(
        self,
        team_id: str,
        channel_id: str,
        chatMessage_id: str,
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error getting channel message: {e}")
            return {"error": str(e)}

    async def send_channel_message(
        self,
        team_id: str,
        channel_id: str,
        data: Dict[str, Any],
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error sending channel message: {e}")
            return {"error": str(e)}

    async def list_team_members(
        self, team_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing team members: {e}")
            return {"error": str(e)}

    async def list_chat_message_replies(
        self, chat_id: str, chatMessage_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing chat message replies: {e}")
            return {"error": str(e)}

    async def reply_to_chat_message(
        self,
        chat_id: str,
        chatMessage_id: str,
        data: Dict[str, Any],
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error replying to chat message: {e}")
            return {"error": str(e)}

    async def get_sharepoint_site_by_path(
        self, site_id: str, path: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get SharePoint site by path."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            # The SDK might have a specific helper for getByPath, or we use the raw way.
            # Usually it's .get_by_path(path)
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
            print(f"Error getting sharepoint site by path: {e}")
            return {"error": str(e)}

    async def get_sharepoint_sites_delta(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting sharepoint sites delta: {e}")
            return {"error": str(e)}

    async def search_query(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Search query."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.search.query.query_post_request_body import (
            QueryPostRequestBody,
        )

        try:
            # This is complex, but basic implementation:
            body = QueryPostRequestBody()
            # body.requests = ...

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
            print(f"Error performing search query: {e}")
            return {"error": str(e)}

    async def list_sharepoint_site_list_items(
        self, site_id: str, list_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing site list items: {e}")
            return {"error": str(e)}

    async def get_sharepoint_site_list_item(
        self,
        site_id: str,
        list_id: str,
        listItem_id: str,
        params: Optional[Dict] = None,
    ) -> Dict[str, Any]:
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
            print(f"Error getting site list item: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Groups
    # =========================================================================

    async def list_groups(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing groups: {e}")
            return {"error": str(e)}

    async def get_group(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting group: {e}")
            return {"error": str(e)}

    async def create_group(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating group: {e}")
            return {"error": str(e)}

    async def update_group(
        self, group_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating group: {e}")
            return {"error": str(e)}

    async def delete_group(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting group: {e}")
            return {"error": str(e)}

    async def list_group_members(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing group members: {e}")
            return {"error": str(e)}

    async def add_group_member(
        self, group_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error adding group member: {e}")
            return {"error": str(e)}

    async def remove_group_member(
        self, group_id: str, member_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error removing group member: {e}")
            return {"error": str(e)}

    async def list_group_owners(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing group owners: {e}")
            return {"error": str(e)}

    async def list_group_conversations(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing group conversations: {e}")
            return {"error": str(e)}

    async def list_group_drives(
        self, group_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing group drives: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Admin / Tenant Management
    # =========================================================================

    async def list_service_health(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List service health overviews."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.admin.service_announcement.health_overviews.to_get_request_configuration()
            )
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
            print(f"Error listing service health: {e}")
            return {"error": str(e)}

    async def get_service_health(
        self, service_name: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service health: {e}")
            return {"error": str(e)}

    async def list_service_health_issues(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List service health issues."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.admin.service_announcement.issues.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.issues.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service health issues: {e}")
            return {"error": str(e)}

    async def get_service_health_issue(
        self, issue_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service health issue: {e}")
            return {"error": str(e)}

    async def list_service_update_messages(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List service update messages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.admin.service_announcement.messages.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.admin.service_announcement.messages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing service update messages: {e}")
            return {"error": str(e)}

    async def get_service_update_message(
        self, message_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting service update message: {e}")
            return {"error": str(e)}

    async def get_admin_sharepoint(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting admin sharepoint: {e}")
            return {"error": str(e)}

    async def update_admin_sharepoint(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating admin sharepoint: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Organization
    # =========================================================================

    async def list_organization(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing organization: {e}")
            return {"error": str(e)}

    async def get_organization(
        self, org_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting organization: {e}")
            return {"error": str(e)}

    async def update_organization(
        self, org_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating organization: {e}")
            return {"error": str(e)}

    async def get_org_branding(
        self, org_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting org branding: {e}")
            return {"error": str(e)}

    async def update_org_branding(
        self, org_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating org branding: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Domains
    # =========================================================================

    async def list_domains(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing domains: {e}")
            return {"error": str(e)}

    async def get_domain(
        self, domain_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting domain: {e}")
            return {"error": str(e)}

    async def create_domain(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating domain: {e}")
            return {"error": str(e)}

    async def delete_domain(
        self, domain_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting domain: {e}")
            return {"error": str(e)}

    async def verify_domain(
        self, domain_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error verifying domain: {e}")
            return {"error": str(e)}

    async def list_domain_service_configuration_records(
        self, domain_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing domain service configuration records: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Subscriptions
    # =========================================================================

    async def list_subscriptions(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing subscriptions: {e}")
            return {"error": str(e)}

    async def get_subscription(
        self, subscription_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting subscription: {e}")
            return {"error": str(e)}

    async def create_subscription(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating subscription: {e}")
            return {"error": str(e)}

    async def update_subscription(
        self, subscription_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating subscription: {e}")
            return {"error": str(e)}

    async def delete_subscription(
        self, subscription_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting subscription: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Communications / Online Meetings
    # =========================================================================

    async def list_online_meetings(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing online meetings: {e}")
            return {"error": str(e)}

    async def get_online_meeting(
        self, meeting_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting online meeting: {e}")
            return {"error": str(e)}

    async def create_online_meeting(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating online meeting: {e}")
            return {"error": str(e)}

    async def update_online_meeting(
        self, meeting_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating online meeting: {e}")
            return {"error": str(e)}

    async def delete_online_meeting(
        self, meeting_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting online meeting: {e}")
            return {"error": str(e)}

    async def list_call_records(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing call records: {e}")
            return {"error": str(e)}

    async def get_call_record(
        self, call_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting call record: {e}")
            return {"error": str(e)}

    async def list_presences(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing presences: {e}")
            return {"error": str(e)}

    async def get_presence(
        self, user_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting presence: {e}")
            return {"error": str(e)}

    async def get_my_presence(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error getting my presence: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Invitations
    # =========================================================================

    async def create_invitation(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating invitation: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Security
    # =========================================================================

    async def list_security_alerts(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing security alerts: {e}")
            return {"error": str(e)}

    async def get_security_alert(
        self, alert_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting security alert: {e}")
            return {"error": str(e)}

    async def update_security_alert(
        self, alert_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating security alert: {e}")
            return {"error": str(e)}

    async def list_security_incidents(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing security incidents: {e}")
            return {"error": str(e)}

    async def get_security_incident(
        self, incident_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting security incident: {e}")
            return {"error": str(e)}

    async def update_security_incident(
        self, incident_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating security incident: {e}")
            return {"error": str(e)}

    async def list_secure_scores(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing secure scores: {e}")
            return {"error": str(e)}

    async def list_threat_intelligence_hosts(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List threat intelligence hosts."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.security.threat_intelligence.hosts.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.threat_intelligence.hosts.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing threat intelligence hosts: {e}")
            return {"error": str(e)}

    async def get_threat_intelligence_host(
        self, host_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting threat intelligence host: {e}")
            return {"error": str(e)}

    async def run_hunting_query(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Run an advanced hunting query."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.security.microsoft_graph_security_run_hunting_query.run_hunting_query_post_request_body import (
            RunHuntingQueryPostRequestBody,
        )

        try:
            body = RunHuntingQueryPostRequestBody()
            body.query = data.get("query")
            request_config = (
                self.client.security.microsoft_graph_security_run_hunting_query.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.security.microsoft_graph_security_run_hunting_query.post(
                body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error running hunting query: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Audit Logs
    # =========================================================================

    async def list_directory_audits(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing directory audits: {e}")
            return {"error": str(e)}

    async def get_directory_audit(
        self, audit_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting directory audit: {e}")
            return {"error": str(e)}

    async def list_sign_in_logs(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing sign-in logs: {e}")
            return {"error": str(e)}

    async def get_sign_in_log(
        self, sign_in_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting sign-in log: {e}")
            return {"error": str(e)}

    async def list_provisioning_logs(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing provisioning logs: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Reports
    # =========================================================================

    async def get_email_activity_report(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting email activity report: {e}")
            return {"error": str(e)}

    async def get_mailbox_usage_report(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting mailbox usage report: {e}")
            return {"error": str(e)}

    async def get_office365_active_users(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting active users report: {e}")
            return {"error": str(e)}

    async def get_sharepoint_activity_report(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting SharePoint activity report: {e}")
            return {"error": str(e)}

    async def get_teams_user_activity(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting Teams user activity report: {e}")
            return {"error": str(e)}

    async def get_onedrive_usage_report(
        self, period: str = "D7", params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"content": native_response.text()}
        except Exception as e:
            print(f"Error getting OneDrive usage report: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Applications
    # =========================================================================

    async def list_applications(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing applications: {e}")
            return {"error": str(e)}

    async def get_application(
        self, app_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting application: {e}")
            return {"error": str(e)}

    async def create_application(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating application: {e}")
            return {"error": str(e)}

    async def update_application(
        self, app_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating application: {e}")
            return {"error": str(e)}

    async def delete_application(
        self, app_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting application: {e}")
            return {"error": str(e)}

    async def add_application_password(
        self, app_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error adding application password: {e}")
            return {"error": str(e)}

    async def remove_application_password(
        self, app_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error removing application password: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Service Principals
    # =========================================================================

    async def list_service_principals(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing service principals: {e}")
            return {"error": str(e)}

    async def get_service_principal(
        self, sp_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting service principal: {e}")
            return {"error": str(e)}

    async def create_service_principal(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating service principal: {e}")
            return {"error": str(e)}

    async def update_service_principal(
        self, sp_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating service principal: {e}")
            return {"error": str(e)}

    async def delete_service_principal(
        self, sp_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting service principal: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Identity (Conditional Access)
    # =========================================================================

    async def list_conditional_access_policies(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List conditional access policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity.conditional_access.policies.to_get_request_configuration()
            )
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
            print(f"Error listing conditional access policies: {e}")
            return {"error": str(e)}

    async def get_conditional_access_policy(
        self, policy_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting conditional access policy: {e}")
            return {"error": str(e)}

    async def create_conditional_access_policy(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            request_config = (
                self.client.identity.conditional_access.policies.to_post_request_configuration()
            )
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
            print(f"Error creating conditional access policy: {e}")
            return {"error": str(e)}

    async def update_conditional_access_policy(
        self, policy_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).patch(
                policy, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating conditional access policy: {e}")
            return {"error": str(e)}

    async def delete_conditional_access_policy(
        self, policy_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).delete(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "deleted"}
        except Exception as e:
            print(f"Error deleting conditional access policy: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Identity Governance
    # =========================================================================

    async def list_access_reviews(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List access review definitions."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_governance.access_reviews.definitions.to_get_request_configuration()
            )
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
            print(f"Error listing access reviews: {e}")
            return {"error": str(e)}

    async def get_access_review(
        self, review_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting access review: {e}")
            return {"error": str(e)}

    async def list_entitlement_access_packages(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List entitlement management access packages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_governance.entitlement_management.access_packages.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_governance.entitlement_management.access_packages.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing access packages: {e}")
            return {"error": str(e)}

    async def list_lifecycle_workflows(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List lifecycle management workflows."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_governance.lifecycle_workflows.workflows.to_get_request_configuration()
            )
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
            print(f"Error listing lifecycle workflows: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Identity Protection
    # =========================================================================

    async def list_risk_detections(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List risk detections."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_protection.risk_detections.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_protection.risk_detections.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing risk detections: {e}")
            return {"error": str(e)}

    async def get_risk_detection(
        self, risk_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting risk detection: {e}")
            return {"error": str(e)}

    async def list_risky_users(self, params: Optional[Dict] = None) -> Dict[str, Any]:
        """List risky users."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.identity_protection.risky_users.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.identity_protection.risky_users.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing risky users: {e}")
            return {"error": str(e)}

    async def get_risky_user(
        self, user_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting risky user: {e}")
            return {"error": str(e)}

    async def dismiss_risky_user(
        self, user_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Dismiss a risky user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.identity_protection.risky_users.dismiss.dismiss_post_request_body import (
            DismissPostRequestBody,
        )

        try:
            body = DismissPostRequestBody()
            body.user_ids = [user_id]
            request_config = (
                self.client.identity_protection.risky_users.dismiss.to_post_request_configuration()
            )
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
            print(f"Error dismissing risky user: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Directory
    # =========================================================================

    async def list_directory_objects(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing directory objects: {e}")
            return {"error": str(e)}

    async def get_directory_object(
        self, object_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting directory object: {e}")
            return {"error": str(e)}

    async def list_directory_roles(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing directory roles: {e}")
            return {"error": str(e)}

    async def get_directory_role(
        self, role_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting directory role: {e}")
            return {"error": str(e)}

    async def list_directory_role_templates(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing directory role templates: {e}")
            return {"error": str(e)}

    async def list_deleted_items(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing deleted items: {e}")
            return {"error": str(e)}

    async def restore_deleted_item(
        self, object_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error restoring deleted item: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Policies
    # =========================================================================

    async def get_authorization_policy(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting authorization policy: {e}")
            return {"error": str(e)}

    async def list_token_lifetime_policies(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List token lifetime policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.policies.token_lifetime_policies.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.token_lifetime_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing token lifetime policies: {e}")
            return {"error": str(e)}

    async def list_token_issuance_policies(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List token issuance policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.policies.token_issuance_policies.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.token_issuance_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing token issuance policies: {e}")
            return {"error": str(e)}

    async def list_permission_grant_policies(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List permission grant policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.policies.permission_grant_policies.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.policies.permission_grant_policies.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing permission grant policies: {e}")
            return {"error": str(e)}

    async def get_admin_consent_policy(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Get the admin consent request policy."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.policies.admin_consent_request_policy.to_get_request_configuration()
            )
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
            print(f"Error getting admin consent policy: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Role Management
    # =========================================================================

    async def list_role_definitions(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List role definitions."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.role_management.directory.role_definitions.to_get_request_configuration()
            )
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
            print(f"Error listing role definitions: {e}")
            return {"error": str(e)}

    async def get_role_definition(
        self, role_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting role definition: {e}")
            return {"error": str(e)}

    async def list_role_assignments(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List role assignments."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.role_management.directory.role_assignments.to_get_request_configuration()
            )
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
            print(f"Error listing role assignments: {e}")
            return {"error": str(e)}

    async def get_role_assignment(
        self, assignment_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting role assignment: {e}")
            return {"error": str(e)}

    async def create_role_assignment(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            request_config = (
                self.client.role_management.directory.role_assignments.to_post_request_configuration()
            )
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
            print(f"Error creating role assignment: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Devices
    # =========================================================================

    async def list_devices(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing devices: {e}")
            return {"error": str(e)}

    async def get_device(
        self, device_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting device: {e}")
            return {"error": str(e)}

    async def delete_device(
        self, device_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting device: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Device Management
    # =========================================================================

    async def list_managed_devices(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List managed devices."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.managed_devices.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.device_management.managed_devices.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing managed devices: {e}")
            return {"error": str(e)}

    async def get_managed_device(
        self, device_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting managed device: {e}")
            return {"error": str(e)}

    async def list_device_compliance_policies(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List device compliance policies."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.device_compliance_policies.to_get_request_configuration()
            )
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
            print(f"Error listing device compliance policies: {e}")
            return {"error": str(e)}

    async def list_device_configurations(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List device configurations."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.device_management.device_configurations.to_get_request_configuration()
            )
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
            print(f"Error listing device configurations: {e}")
            return {"error": str(e)}

    async def wipe_managed_device(
        self, device_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).wipe.post(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "wipe initiated"}
        except Exception as e:
            print(f"Error wiping managed device: {e}")
            return {"error": str(e)}

    async def retire_managed_device(
        self, device_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).retire.post(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return {"status": "retire initiated"}
        except Exception as e:
            print(f"Error retiring managed device: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Education
    # =========================================================================

    async def list_education_classes(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing education classes: {e}")
            return {"error": str(e)}

    async def get_education_class(
        self, class_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting education class: {e}")
            return {"error": str(e)}

    async def list_education_schools(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing education schools: {e}")
            return {"error": str(e)}

    async def get_education_school(
        self, school_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting education school: {e}")
            return {"error": str(e)}

    async def list_education_users(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing education users: {e}")
            return {"error": str(e)}

    async def list_education_assignments(
        self, class_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing education assignments: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Agreements
    # =========================================================================

    async def list_agreements(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing agreements: {e}")
            return {"error": str(e)}

    async def get_agreement(
        self, agreement_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting agreement: {e}")
            return {"error": str(e)}

    async def create_agreement(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating agreement: {e}")
            return {"error": str(e)}

    async def delete_agreement(
        self, agreement_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting agreement: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Places
    # =========================================================================

    async def list_rooms(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing rooms: {e}")
            return {"error": str(e)}

    async def list_room_lists(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing room lists: {e}")
            return {"error": str(e)}

    async def get_place(
        self, place_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting place: {e}")
            return {"error": str(e)}

    async def update_place(
        self, place_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error updating place: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Print
    # =========================================================================

    async def list_printers(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing printers: {e}")
            return {"error": str(e)}

    async def get_printer(
        self, printer_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting printer: {e}")
            return {"error": str(e)}

    async def list_print_jobs(
        self, printer_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing print jobs: {e}")
            return {"error": str(e)}

    async def create_print_job(
        self, printer_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating print job: {e}")
            return {"error": str(e)}

    async def list_print_shares(self, params: Optional[Dict] = None) -> Dict[str, Any]:
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
            print(f"Error listing print shares: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Privacy
    # =========================================================================

    async def list_subject_rights_requests(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List subject rights requests."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.privacy.subject_rights_requests.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.privacy.subject_rights_requests.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing subject rights requests: {e}")
            return {"error": str(e)}

    async def get_subject_rights_request(
        self, request_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting subject rights request: {e}")
            return {"error": str(e)}

    async def create_subject_rights_request(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create a subject rights request."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.subject_rights_request import SubjectRightsRequest

        try:
            srr = SubjectRightsRequest()
            srr.display_name = data.get("displayName")
            srr.description = data.get("description")
            request_config = (
                self.client.privacy.subject_rights_requests.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.privacy.subject_rights_requests.post(
                srr, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating subject rights request: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Solutions (Bookings & Virtual Events)
    # =========================================================================

    async def list_booking_businesses(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing booking businesses: {e}")
            return {"error": str(e)}

    async def get_booking_business(
        self, business_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting booking business: {e}")
            return {"error": str(e)}

    async def list_booking_appointments(
        self, business_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing booking appointments: {e}")
            return {"error": str(e)}

    async def create_booking_appointment(
        self, business_id: str, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating booking appointment: {e}")
            return {"error": str(e)}

    async def list_virtual_events(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List virtual event townhalls."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.solutions.virtual_events.townhalls.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.solutions.virtual_events.townhalls.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing virtual events: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Storage
    # =========================================================================

    async def list_file_storage_containers(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List file storage containers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.storage.file_storage.containers.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.storage.file_storage.containers.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing file storage containers: {e}")
            return {"error": str(e)}

    async def get_file_storage_container(
        self, container_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting file storage container: {e}")
            return {"error": str(e)}

    async def create_file_storage_container(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """Create a file storage container."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.file_storage_container import FileStorageContainer

        try:
            container = FileStorageContainer()
            container.display_name = data.get("displayName")
            container.container_type_id = data.get("containerTypeId")
            request_config = (
                self.client.storage.file_storage.containers.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.storage.file_storage.containers.post(
                container, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating file storage container: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Employee Experience
    # =========================================================================

    async def list_learning_providers(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List learning providers."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.employee_experience.learning_providers.to_get_request_configuration()
            )
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
            print(f"Error listing learning providers: {e}")
            return {"error": str(e)}

    async def get_learning_provider(
        self, provider_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting learning provider: {e}")
            return {"error": str(e)}

    async def list_learning_course_activities(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List learning course activities for the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.me.employee_experience.learning_course_activities.to_get_request_configuration()
            )
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
            print(f"Error listing learning course activities: {e}")
            return {"error": str(e)}

    # =========================================================================
    # External Connectors
    # =========================================================================

    async def list_external_connections(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error listing external connections: {e}")
            return {"error": str(e)}

    async def get_external_connection(
        self, connection_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error getting external connection: {e}")
            return {"error": str(e)}

    async def create_external_connection(
        self, data: Dict[str, Any], params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error creating external connection: {e}")
            return {"error": str(e)}

    async def delete_external_connection(
        self, connection_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            print(f"Error deleting external connection: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Information Protection
    # =========================================================================

    async def list_sensitivity_labels(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List sensitivity labels."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.information_protection.policy.labels.to_get_request_configuration()
            )
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
            print(f"Error listing sensitivity labels: {e}")
            return {"error": str(e)}

    async def get_sensitivity_label(
        self, label_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting sensitivity label: {e}")
            return {"error": str(e)}

    # =========================================================================
    # Tenant Relationships
    # =========================================================================

    async def list_delegated_admin_relationships(
        self, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
        """List delegated admin relationships."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.tenant_relationships.delegated_admin_relationships.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await self.client.tenant_relationships.delegated_admin_relationships.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing delegated admin relationships: {e}")
            return {"error": str(e)}

    async def get_delegated_admin_relationship(
        self, rel_id: str, params: Optional[Dict] = None
    ) -> Dict[str, Any]:
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
            ).get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting delegated admin relationship: {e}")
            return {"error": str(e)}
