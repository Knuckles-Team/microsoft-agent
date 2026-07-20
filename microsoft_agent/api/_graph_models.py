"""Strict conversion helpers for generated Microsoft Graph SDK models."""

from __future__ import annotations

import base64
import binascii
import json
import re
from collections.abc import Callable
from typing import Any, TypeVar
from urllib.parse import urlparse

from kiota_abstractions.serialization import Parsable
from kiota_serialization_json.json_parse_node_factory import JsonParseNodeFactory

GraphModel = TypeVar("GraphModel", bound=Parsable)


def graph_model_from_dict(
    data: dict[str, Any], factory: Callable[[], GraphModel]
) -> GraphModel:
    """Deserialize Graph JSON into a complete generated SDK model."""

    parse_node = JsonParseNodeFactory().get_root_parse_node(
        "application/json", json.dumps(data).encode("utf-8")
    )
    return parse_node.get_object_value(factory)


def decode_graph_base64(value: Any, field_name: str) -> bytes:
    if not isinstance(value, str) or not value:
        raise ValueError(f"{field_name} must be a non-empty base64 string")
    try:
        return base64.b64decode(value, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError(f"{field_name} must be valid base64") from exc


def chat_message_from_dict(data: dict[str, Any]) -> Any:
    """Build a complete Teams message, including validated hosted content."""

    from msgraph.generated.models.chat_message import ChatMessage

    body_data = data.get("body")
    if not isinstance(body_data, dict):
        raise ValueError("Message body content is required")
    content = body_data.get("content")
    if not isinstance(content, str) or not content:
        raise ValueError("Message body content is required")

    message = graph_model_from_dict(data, ChatMessage)
    hosted_data = data.get("hostedContents")
    if hosted_data is not None:
        if not isinstance(hosted_data, list):
            raise ValueError("hostedContents must be a list")
        hosted_contents = message.hosted_contents or []
        if len(hosted_contents) != len(hosted_data):
            raise ValueError("hostedContents could not be parsed")
        for index, (hosted, source) in enumerate(
            zip(hosted_contents, hosted_data, strict=True)
        ):
            if not isinstance(source, dict):
                raise ValueError("Each hostedContents entry must be an object")
            if "contentBytes" in source:
                hosted.content_bytes = decode_graph_base64(
                    source["contentBytes"], f"hostedContents[{index}].contentBytes"
                )
    return message


def validated_planner_etag(etag: str | None, params: dict[str, Any] | None) -> str:
    """Return a bounded Planner entity tag for the required If-Match header."""

    candidate: Any = etag
    if candidate is None and params:
        candidate = params.get("If-Match", params.get("ifMatch", params.get("etag")))
    if not isinstance(candidate, str):
        raise ValueError("A Planner ETag is required for the If-Match header")
    candidate = candidate.strip()
    if len(candidate) > 1024 or not re.fullmatch(r'(?:W/)?"[^"\r\n]+"', candidate):
        raise ValueError("Planner ETag must be a quoted HTTP entity tag")
    return candidate


def comma_separated_values(value: Any, field_name: str) -> list[str]:
    """Validate and split a bounded comma-separated Graph query value."""

    if not isinstance(value, str) or not value or len(value) > 4096:
        raise ValueError(f"{field_name} must be a non-empty string")
    values = [item.strip() for item in value.split(",")]
    if any(not item for item in values):
        raise ValueError(f"{field_name} contains an empty value")
    return values


def validated_sharepoint_delta_url(value: Any) -> str:
    """Accept only opaque continuation links for the fixed Graph delta route."""

    if not isinstance(value, str) or not value or len(value) > 32768:
        raise ValueError("SharePoint delta continuation URL is invalid")
    parsed = urlparse(value)
    try:
        port = parsed.port
    except ValueError as exc:
        raise ValueError("SharePoint delta continuation URL is invalid") from exc
    if (
        parsed.scheme != "https"
        or (parsed.hostname or "").casefold() != "graph.microsoft.com"
        or port not in {None, 443}
        or parsed.username
        or parsed.password
        or parsed.fragment
        or parsed.path.rstrip("/") != "/v1.0/sites/delta"
    ):
        raise ValueError(
            "SharePoint delta continuation URL must target the Microsoft Graph "
            "v1.0 sites/delta endpoint"
        )
    return value
