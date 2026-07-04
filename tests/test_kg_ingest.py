"""Native epistemic-graph typed-node ingestion — Wire-First coverage.

Exercises the real ``ingest_entities`` / ``ingest_messages`` / ``ingest_events`` /
``ingest_files`` / ``ingest_users`` seam with a fake engine client (no engine
required), asserting the txn add_node/commit + edge calls and the Microsoft Graph
record → typed-node mapping. CONCEPT:AU-KG.ingest.enterprise-source-extractor.
"""

from __future__ import annotations

from microsoft_agent.kg_ingest import (
    _records,
    ingest_entities,
    ingest_events,
    ingest_files,
    ingest_messages,
    ingest_users,
)


class _FakeTxn:
    def __init__(self):
        self.nodes = {}
        self.committed = False

    def begin(self, graph=None):
        self.graph = graph
        return "txn-1"

    def add_node(self, txn, node_id, props):
        self.nodes[node_id] = props

    def commit(self, txn):
        self.committed = True
        return True


class _FakeEdges:
    def __init__(self):
        self.edges = []

    def add(self, src, dst, props):
        self.edges.append((src, dst, props))


class _FakeClient:
    def __init__(self):
        self.txn = _FakeTxn()
        self.edges = _FakeEdges()


def test_ingest_entities_writes_nodes_and_edges():
    c = _FakeClient()
    res = ingest_entities(
        [
            {"id": "a", "type": "Message", "subject": "hi"},
            {"id": "b", "type": "Person"},
        ],
        [{"source": "a", "target": "b", "type": "sentBy"}],
        client=c,
        graph="__commons__",
    )
    assert res == {"nodes": 2, "edges": 1}
    assert c.txn.committed is True
    assert set(c.txn.nodes) == {"a", "b"}
    # provenance is stamped
    assert c.txn.nodes["a"]["source"] == "microsoft-agent"
    assert c.txn.nodes["a"]["domain"] == "microsoft"
    assert c.edges.edges == [("a", "b", {"type": "sentBy"})]


def test_ingest_messages_maps_message_person_and_body():
    c = _FakeClient()
    res = ingest_messages(
        [
            {
                "id": "MSG1",
                "subject": "Quarterly report",
                "bodyPreview": "see attached",
                "body": {"content": "Full body text", "contentType": "text"},
                "webLink": "https://outlook/MSG1",
                "receivedDateTime": "2026-07-01T10:00:00Z",
                "from": {"emailAddress": {"address": "ana@contoso.com", "name": "Ana"}},
                "toRecipients": [
                    {"emailAddress": {"address": "bo@contoso.com", "name": "Bo"}}
                ],
            }
        ],
        client=c,
        graph="__commons__",
    )
    assert res is not None
    msg = c.txn.nodes["microsoft:Message:MSG1"]
    assert msg["type"] == "Message"
    assert msg["subject"] == "Quarterly report"
    assert msg["externalToolId"] == "MSG1"
    # sender + recipient became :Person nodes
    assert c.txn.nodes["microsoft:Person:ana@contoso.com"]["type"] == "Person"
    assert c.txn.nodes["microsoft:Person:bo@contoso.com"]["type"] == "Person"
    # body :Document written (second commit) + links present
    assert c.txn.nodes["microsoft:Document:message:MSG1"]["type"] == "Document"
    edge_types = {e[2]["type"] for e in c.edges.edges}
    assert {"sentBy", "sentTo", "hasBody"} <= edge_types


def test_ingest_events_maps_event_and_attendees():
    c = _FakeClient()
    ingest_events(
        [
            {
                "id": "EV1",
                "subject": "Design sync",
                "start": {"dateTime": "2026-07-05T15:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-07-05T15:30:00", "timeZone": "UTC"},
                "organizer": {"emailAddress": {"address": "ana@contoso.com"}},
                "attendees": [
                    {"emailAddress": {"address": "bo@contoso.com"}, "type": "required"}
                ],
            }
        ],
        client=c,
        graph="__commons__",
    )
    ev = c.txn.nodes["microsoft:Event:EV1"]
    assert ev["type"] == "Event"
    assert ev["startDateTime"] == "2026-07-05T15:00:00"
    edge_types = {e[2]["type"] for e in c.edges.edges}
    assert {"organizedBy", "hasAttendee"} <= edge_types


def test_ingest_files_maps_file_and_owner():
    c = _FakeClient()
    ingest_files(
        [
            {
                "id": "F1",
                "name": "budget.xlsx",
                "webUrl": "https://sp/F1",
                "size": 2048,
                "file": {"mimeType": "application/vnd.ms-excel"},
                "lastModifiedDateTime": "2026-06-30T09:00:00Z",
                "createdBy": {"user": {"id": "u9", "displayName": "Cy"}},
            }
        ],
        client=c,
        graph="__commons__",
    )
    f = c.txn.nodes["microsoft:File:F1"]
    assert f["type"] == "File"
    assert f["mimeType"] == "application/vnd.ms-excel"
    assert f["sizeBytes"] == 2048
    assert c.txn.nodes["microsoft:Person:u9"]["type"] == "Person"
    assert c.edges.edges[0][2]["type"] == "ownedByPerson"


def test_ingest_users_maps_person():
    c = _FakeClient()
    ingest_users(
        [
            {
                "id": "u1",
                "displayName": "Ana Admin",
                "mail": "ana@contoso.com",
                "userPrincipalName": "ana@contoso.com",
                "jobTitle": "Ops",
            }
        ],
        client=c,
        graph="__commons__",
    )
    p = c.txn.nodes["microsoft:Person:u1"]
    assert p["type"] == "Person"
    assert p["emailAddress"] == "ana@contoso.com"
    assert p["jobTitle"] == "Ops"


def test_records_normalizes_graph_value_envelope():
    assert _records({"value": [{"id": "a"}, {"id": "b"}]}) == [
        {"id": "a"},
        {"id": "b"},
    ]
    assert _records({"error": "boom"}) == []
    assert _records({"id": "solo"}) == [{"id": "solo"}]
    assert _records(None) == []


def test_ingest_noops_without_engine():
    # No injected client + no reachable engine -> clean no-op.
    assert ingest_entities([{"id": "a", "type": "Message"}]) is None


def test_ingest_empty_is_noop():
    assert ingest_entities([], client=_FakeClient()) is None
    assert ingest_messages([], client=_FakeClient()) is None
    assert ingest_events([], client=_FakeClient()) is None
