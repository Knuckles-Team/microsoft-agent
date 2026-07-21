"""Tests for the keyed zero-PII Microsoft source projection."""

from __future__ import annotations

import pytest

from microsoft_agent import kg_ingest


class ProjectionSettings:
    key = b"test-only-pseudonymization-key-32-bytes"

    def require_ingestion_pseudonymization_key(self) -> bytes:
        return self.key


@pytest.fixture(autouse=True)
def projection_settings(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(kg_ingest, "get_settings", lambda: ProjectionSettings())


def test_message_projection_exposes_only_opaque_structure() -> None:
    source_record = {
        "id": "provider-message-id",
        "subject": "private subject",
        "bodyPreview": "private preview",
        "body": {"content": "private body", "contentType": "text"},
        "webLink": "https://provider.example.invalid/private-message",
        "receivedDateTime": "2026-07-01T10:00:00Z",
        "from": {
            "emailAddress": {
                "address": "sender@example.invalid",
                "name": "Sender Name",
            }
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "address": "recipient@example.invalid",
                    "name": "Recipient Name",
                }
            }
        ],
    }

    projection = kg_ingest.project_records("messages", [source_record])

    assert {node["node_type"] for node in projection["records"]} == {
        "Message",
        "Person",
    }
    assert {edge["relationship"] for edge in projection["relationships"]} == {
        "sentBy",
        "sentTo",
    }
    serialized = repr(projection)
    for raw_value in (
        "provider-message-id",
        "private subject",
        "private preview",
        "private body",
        "private-message",
        "sender@example.invalid",
        "recipient@example.invalid",
        "Sender Name",
        "Recipient Name",
    ):
        assert raw_value not in serialized


def test_event_projection_exposes_only_opaque_structure() -> None:
    projection = kg_ingest.project_records(
        "events",
        [
            {
                "id": "provider-event-id",
                "subject": "private event",
                "organizer": {"emailAddress": {"address": "owner@example.invalid"}},
                "attendees": [{"emailAddress": {"address": "guest@example.invalid"}}],
            }
        ],
    )

    assert {tuple(sorted(node)) for node in projection["records"]} == {
        ("id", "node_type")
    }
    assert {edge["relationship"] for edge in projection["relationships"]} == {
        "organizedBy",
        "hasAttendee",
    }
    assert "provider-event-id" not in repr(projection)
    assert "owner@example.invalid" not in repr(projection)
    assert "guest@example.invalid" not in repr(projection)


def test_file_and_user_projections_discard_provider_attributes() -> None:
    files = kg_ingest.project_records(
        "files",
        [
            {
                "id": "provider-file-id",
                "name": "private-budget.xlsx",
                "webUrl": "https://provider.example.invalid/private-file",
                "createdBy": {
                    "user": {"id": "provider-person-id", "displayName": "Owner"}
                },
            }
        ],
    )
    users = kg_ingest.project_records(
        "users",
        [
            {
                "id": "provider-person-id",
                "displayName": "Private Name",
                "userPrincipalName": "operator@example.invalid",
            }
        ],
    )

    assert files["relationships"][0]["relationship"] == "ownedByPerson"
    assert users["records"][0]["node_type"] == "Person"
    assert set(users["records"][0]) == {"id", "node_type"}
    serialized = repr((files, users))
    assert "provider-file-id" not in serialized
    assert "private-budget.xlsx" not in serialized
    assert "provider-person-id" not in serialized
    assert "operator@example.invalid" not in serialized


def test_opaque_ids_are_keyed_and_domain_separated() -> None:
    key = ProjectionSettings.key
    first = kg_ingest._opaque_id("Person", "same-provider-id", key)
    second = kg_ingest._opaque_id("Person", "same-provider-id", key)
    other_type = kg_ingest._opaque_id("Message", "same-provider-id", key)
    other_key = kg_ingest._opaque_id("Person", "same-provider-id", b"x" * 32)

    assert first == second
    assert first != other_type
    assert first != other_key
    assert "same-provider-id" not in first


def test_missing_pseudonymization_key_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class MissingKey:
        def require_ingestion_pseudonymization_key(self) -> bytes:
            raise ValueError("pseudonymization key required")

    monkeypatch.setattr(kg_ingest, "get_settings", lambda: MissingKey())
    with pytest.raises(ValueError, match="pseudonymization key required"):
        kg_ingest.project_records("messages", [{"id": "provider-message-id"}])


def test_empty_projection_is_an_explicit_zero_change() -> None:
    assert kg_ingest.project_records("messages", []) == {
        "records": [],
        "relationships": [],
    }


def test_records_normalizes_graph_value_envelope() -> None:
    assert kg_ingest._records({"value": [{"id": "a"}, {"id": "b"}]}) == [
        {"id": "a"},
        {"id": "b"},
    ]
    with pytest.raises(ValueError, match="provider returned an error"):
        kg_ingest._records({"error": "provider-error"})
    assert kg_ingest._records({"id": "single"}) == [{"id": "single"}]
    assert kg_ingest._records(None) == []
