from typing import Any

from microsoft_agent.api._graph_models import (
    chat_message_from_dict,
    graph_model_from_dict,
)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def send_mail(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send mail."""
        from msgraph.generated.models.message import Message
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        try:
            request_body = SendMailPostRequestBody()
            request_body.message = graph_model_from_dict(
                data.get("message", {}), Message
            )
            request_body.save_to_sent_items = data.get("saveToSentItems", True)

            await self.client.me.send_mail.post(request_body)
            return {"status": "success"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_draft_email(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create draft email."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.message import Message

        try:
            message = graph_model_from_dict(data, Message)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def delete_mail_message(
        self, message_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete a message."""
        try:
            await self.client.me.messages.by_message_id(message_id).delete()
            return {"status": "success"}
        except Exception:
            return {"error": "Operation failed"}

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
            request_body = graph_model_from_dict(data, MovePostRequestBody)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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

        message = graph_model_from_dict(data, Message)

        request_config = MessageItemRequestBuilder.MessageItemRequestBuilderPatchRequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())]
        )

        try:
            native_response = await self.client.me.messages.by_message_id(
                message_id
            ).patch(message, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception:
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def send_shared_mailbox_mail(
        self, user_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send mail from a shared mailbox."""
        from msgraph.generated.models.message import Message
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
            SendMailPostRequestBody,
        )

        try:
            request_body = SendMailPostRequestBody()
            request_body.message = graph_model_from_dict(
                data.get("message", {}), Message
            )
            request_body.save_to_sent_items = data.get("saveToSentItems", True)

            await self.client.users.by_user_id(user_id).send_mail.post(request_body)

            return {"status": "success"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def send_chat_message(
        self, chat_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Send chat message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            message = chat_message_from_dict(data)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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

        try:
            message = chat_message_from_dict(data)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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

        try:
            message = chat_message_from_dict(data)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_channel_message_replies(
        self,
        team_id: str,
        channel_id: str,
        chatMessage_id: str,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """List replies to a Teams channel message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            replies = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies
            )
            request_config = replies.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await replies.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def reply_to_channel_message(
        self,
        team_id: str,
        channel_id: str,
        chatMessage_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Reply to a Teams channel message."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            message = chat_message_from_dict(data)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
            replies = (
                self.client.teams.by_team_id(team_id)
                .channels.by_channel_id(channel_id)
                .messages.by_chat_message_id(chatMessage_id)
                .replies
            )
            request_config = replies.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )
            native_response = await replies.post(
                message, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}
