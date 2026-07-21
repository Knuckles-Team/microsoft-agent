"""Keyed zero-PII Microsoft Graph source projection.
Microsoft records are reduced in memory to opaque typed nodes and structural
relationships. Names, addresses, subjects, bodies, filenames, URLs, timestamps,
attachment bytes, and provider identifiers never cross the projection boundary.
GraphOS owns validation, quarantine, and ChangeEnvelope materialization through
the signed source-connector manifest; this provider exposes no graph-write plane.
"""

from __future__ import annotations

import hashlib
import hmac
from collections.abc import Callable
from typing import Any

from microsoft_agent.settings import get_settings

Projection = tuple[list[dict[str, Any]], list[dict[str, str]]]


def _records(response: Any) -> list[dict[str, Any]]:
    """Normalize a Graph collection response and reject provider errors."""

    if response is None:
        return []
    data = response
    if isinstance(response, dict):
        if response.get("error"):
            raise ValueError("Microsoft provider returned an error response")
        if "value" in response:
            data = response["value"]
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError("Microsoft provider returned an invalid record collection")
    records = [record for record in data if isinstance(record, dict)]
    if len(records) != len(data):
        raise ValueError("Microsoft provider returned a malformed record")
    return records


def _opaque_id(node_type: str, provider_id: Any, key: bytes) -> str:
    """Return a domain-separated keyed identifier without retaining source PII."""

    digest = hmac.new(
        key,
        f"microsoft-agent\x00{node_type}\x00{provider_id}".encode(),
        hashlib.sha256,
    ).hexdigest()
    return f"microsoft:{node_type}:{digest}"


def _projection_key() -> bytes:
    return get_settings().require_ingestion_pseudonymization_key()


def _person_node(user: dict[str, Any] | None, key: bytes) -> dict[str, str] | None:
    if not user:
        return None
    provider_id = user.get("id") or user.get("address") or user.get("userPrincipalName")
    if not provider_id:
        return None
    return {
        "id": _opaque_id("Person", provider_id, key),
        "node_type": "Person",
    }


def _append_once(
    entities: list[dict[str, Any]],
    seen: set[str],
    node: dict[str, Any] | None,
) -> None:
    if node is not None and node["id"] not in seen:
        entities.append(node)
        seen.add(str(node["id"]))


def _project_messages(records: list[dict[str, Any]], key: bytes) -> Projection:
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, str]] = []
    seen: set[str] = set()
    for record in records:
        provider_id = record.get("id")
        if not provider_id:
            raise ValueError("Microsoft message record has no identifier")
        node_id = _opaque_id("Message", provider_id, key)
        _append_once(entities, seen, {"id": node_id, "node_type": "Message"})

        sender = record.get("from") or record.get("sender") or {}
        if isinstance(sender, dict):
            sender = sender.get("emailAddress") or sender.get("user") or sender
        sender_node = _person_node(sender if isinstance(sender, dict) else None, key)
        _append_once(entities, seen, sender_node)
        if sender_node:
            relationships.append(
                {
                    "source": node_id,
                    "target": sender_node["id"],
                    "relationship": "sentBy",
                }
            )

        for recipient in record.get("toRecipients") or []:
            address = (
                recipient.get("emailAddress") if isinstance(recipient, dict) else None
            )
            recipient_node = _person_node(address, key)
            _append_once(entities, seen, recipient_node)
            if recipient_node:
                relationships.append(
                    {
                        "source": node_id,
                        "target": recipient_node["id"],
                        "relationship": "sentTo",
                    }
                )
    return entities, relationships


def _project_events(records: list[dict[str, Any]], key: bytes) -> Projection:
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, str]] = []
    seen: set[str] = set()
    for record in records:
        provider_id = record.get("id")
        if not provider_id:
            raise ValueError("Microsoft event record has no identifier")
        node_id = _opaque_id("Event", provider_id, key)
        _append_once(entities, seen, {"id": node_id, "node_type": "Event"})

        organizer = (record.get("organizer") or {}).get("emailAddress")
        organizer_node = _person_node(organizer, key)
        _append_once(entities, seen, organizer_node)
        if organizer_node:
            relationships.append(
                {
                    "source": node_id,
                    "target": organizer_node["id"],
                    "relationship": "organizedBy",
                }
            )

        for attendee in record.get("attendees") or []:
            address = (
                attendee.get("emailAddress") if isinstance(attendee, dict) else None
            )
            attendee_node = _person_node(address, key)
            _append_once(entities, seen, attendee_node)
            if attendee_node:
                relationships.append(
                    {
                        "source": node_id,
                        "target": attendee_node["id"],
                        "relationship": "hasAttendee",
                    }
                )
    return entities, relationships


def _project_files(records: list[dict[str, Any]], key: bytes) -> Projection:
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, str]] = []
    seen: set[str] = set()
    for record in records:
        provider_id = record.get("id")
        if not provider_id:
            raise ValueError("Microsoft file record has no identifier")
        node_id = _opaque_id("File", provider_id, key)
        _append_once(entities, seen, {"id": node_id, "node_type": "File"})
        owner = (record.get("createdBy") or {}).get("user")
        owner_node = _person_node(owner, key)
        _append_once(entities, seen, owner_node)
        if owner_node:
            relationships.append(
                {
                    "source": node_id,
                    "target": owner_node["id"],
                    "relationship": "ownedByPerson",
                }
            )
    return entities, relationships


def _project_users(records: list[dict[str, Any]], key: bytes) -> Projection:
    entities: list[dict[str, Any]] = []
    seen: set[str] = set()
    for record in records:
        node = _person_node(record, key)
        if node is None:
            addresses = record.get("emailAddresses") or []
            address = (
                addresses[0].get("address")
                if addresses and isinstance(addresses[0], dict)
                else None
            )
            node = _person_node({"id": record.get("id"), "address": address}, key)
        if node is None:
            raise ValueError("Microsoft user record has no identifier")
        _append_once(entities, seen, node)
    return entities, []


_PROJECTORS: dict[str, Callable[[list[dict[str, Any]], bytes], Projection]] = {
    "messages": _project_messages,
    "events": _project_events,
    "files": _project_files,
    "users": _project_users,
}


def project_records(kind: str, records: list[dict[str, Any]]) -> dict[str, Any]:
    """Return the signed source-connector projection with no raw provider values."""

    projector = _PROJECTORS.get(kind)
    if projector is None:
        raise ValueError("Unknown Microsoft projection kind")
    if not records:
        return {"records": [], "relationships": []}
    entities, relationships = projector(records, _projection_key())
    return {"records": entities, "relationships": relationships}
