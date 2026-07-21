"""Regression tests for Teams message send and channel reply operations."""

from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock

import pytest

from microsoft_agent.api_client import MicrosoftGraphApi


def api_with_client(client: MagicMock) -> MicrosoftGraphApi:
    api = object.__new__(MicrosoftGraphApi)
    api.client = client
    return api


def native_response(payload: dict) -> MagicMock:
    response = MagicMock()
    response.raise_for_status.return_value = None
    response.json.return_value = payload
    return response


@pytest.mark.asyncio
async def test_send_chat_message_sets_html_content_type() -> None:
    client = MagicMock()
    messages = client.chats.by_chat_id.return_value.messages
    messages.to_post_request_configuration.return_value = SimpleNamespace(options=[])
    messages.post = AsyncMock(return_value=native_response({"id": "message-1"}))
    api = api_with_client(client)

    result = await api.send_chat_message(
        "chat-1",
        {"body": {"content": "<b>Hello</b>", "contentType": "html"}},
    )

    assert result == {"id": "message-1"}
    awaited = messages.post.await_args
    assert awaited is not None
    message = awaited.args[0]
    assert message.body.content == "<b>Hello</b>"
    assert message.body.content_type.value.casefold() == "html"


@pytest.mark.asyncio
async def test_send_channel_message_rejects_missing_content_without_network() -> None:
    client = MagicMock()
    api = api_with_client(client)

    result = await api.send_channel_message("team-1", "channel-1", {"body": {}})

    assert result == {"error": "Message body content is required"}
    client.teams.by_team_id.return_value.channels.by_channel_id.return_value.messages.post.assert_not_called()


@pytest.mark.asyncio
async def test_channel_reply_uses_team_channel_message_replies_endpoint() -> None:
    client = MagicMock()
    replies = client.teams.by_team_id.return_value.channels.by_channel_id.return_value.messages.by_chat_message_id.return_value.replies
    replies.to_post_request_configuration.return_value = SimpleNamespace(options=[])
    replies.post = AsyncMock(return_value=native_response({"id": "reply-1"}))
    api = api_with_client(client)

    result = await api.reply_to_channel_message(
        "team-1",
        "channel-1",
        "message-1",
        {"body": {"content": "Reply", "contentType": "text"}},
    )

    assert result == {"id": "reply-1"}
    client.teams.by_team_id.assert_called_with("team-1")
    replies.post.assert_awaited_once()


@pytest.mark.asyncio
async def test_teams_message_mutations_preserve_complete_payload() -> None:
    """Chat and channel sends/replies retain rich Graph message fields."""
    client = MagicMock()
    chat_messages = client.chats.by_chat_id.return_value.messages
    chat_replies = chat_messages.by_chat_message_id.return_value.replies
    channel_messages = client.teams.by_team_id.return_value.channels.by_channel_id.return_value.messages
    channel_replies = channel_messages.by_chat_message_id.return_value.replies
    for builder in (chat_messages, chat_replies, channel_messages, channel_replies):
        builder.to_post_request_configuration.return_value = SimpleNamespace(options=[])
        builder.post = AsyncMock(return_value=native_response({"id": "message-1"}))

    data = {
        "body": {
            "contentType": "html",
            "content": 'Hello <at id="0">Ada</at><img src="../hostedContents/1/$value">',
        },
        "subject": "Release review",
        "importance": "high",
        "mentions": [
            {
                "id": 0,
                "mentionText": "Ada",
                "mentioned": {
                    "user": {
                        "id": "user-1",
                        "displayName": "Ada Lovelace",
                        "userIdentityType": "aadUser",
                    }
                },
            }
        ],
        "attachments": [
            {
                "id": "attachment-1",
                "contentType": "reference",
                "contentUrl": "https://contoso.sharepoint.com/release.docx",
                "name": "release.docx",
            }
        ],
        "hostedContents": [
            {
                "@microsoft.graph.temporaryId": "1",
                "contentBytes": "aGVsbG8=",
                "contentType": "image/png",
            }
        ],
    }
    api = api_with_client(client)

    assert await api.send_chat_message("chat-1", data) == {"id": "message-1"}
    assert await api.reply_to_chat_message("chat-1", "message-1", data) == {
        "id": "message-1"
    }
    assert await api.send_channel_message("team-1", "channel-1", data) == {
        "id": "message-1"
    }
    assert await api.reply_to_channel_message(
        "team-1", "channel-1", "message-1", data
    ) == {"id": "message-1"}

    for builder in (chat_messages, chat_replies, channel_messages, channel_replies):
        awaited_call = builder.post.await_args
        assert awaited_call is not None
        message = awaited_call.args[0]
        assert message.body.content_type.value == "html"
        assert message.subject == "Release review"
        assert message.importance.value == "high"
        assert message.mentions[0].mention_text == "Ada"
        assert message.mentions[0].mentioned.user.id == "user-1"
        assert message.attachments[0].name == "release.docx"
        assert message.attachments[0].content_type == "reference"
        assert message.hosted_contents[0].content_bytes == b"hello"
        assert message.hosted_contents[0].content_type == "image/png"
        assert (
            message.hosted_contents[0].additional_data["@microsoft.graph.temporaryId"]
            == "1"
        )
