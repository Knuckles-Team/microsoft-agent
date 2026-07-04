"""Native epistemic-graph blob ingestion — Wire-First live-path coverage.

Exercises the real ``ingest_attachment`` seam with a fake ``MediaStore`` (no engine
required) and asserts the base64 attachment bytes + metadata reach ``store_media``.
CONCEPT:AU-KG.ingest.list-durable-media.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass

from microsoft_agent.kg_media import ingest_attachment


@dataclass
class _Stored:
    asset_id: str
    digest: str


class _FakeMediaStore:
    def __init__(self):
        self.calls = []

    def store_media(self, data, **kw):
        self.calls.append((data, kw))
        return _Stored(asset_id="media:deadbeef", digest="deadbeef")


def test_ingest_attachment_stores_bytes_and_metadata():
    payload = b"\x00\x01pdf-bytes\x02"
    att = {
        "id": "ATT1",
        "name": "report.pdf",
        "contentType": "application/pdf",
        "contentBytes": base64.b64encode(payload).decode(),
    }
    store = _FakeMediaStore()
    res = ingest_attachment(att, message_id="MSG1", media_store=store)

    assert res is not None
    assert res["asset_id"] == "media:deadbeef"
    assert res["media_type"] == "file"
    assert res["size_bytes"] == len(payload)

    assert len(store.calls) == 1
    data, kw = store.calls[0]
    assert data == payload
    assert kw["mime_type"] == "application/pdf"
    assert kw["name"] == "report.pdf"
    assert kw["source"] == "microsoft-agent"
    assert kw["extra"]["message_id"] == "MSG1"
    assert kw["extra"]["attachment_id"] == "ATT1"


def test_ingest_attachment_image_media_type():
    att = {
        "id": "ATT2",
        "name": "pic.png",
        "contentType": "image/png",
        "contentBytes": base64.b64encode(b"img").decode(),
    }
    store = _FakeMediaStore()
    res = ingest_attachment(att, media_store=store)
    assert res["media_type"] == "image"


def test_ingest_attachment_noops_without_bytes():
    assert (
        ingest_attachment({"id": "x", "name": "n"}, media_store=_FakeMediaStore())
        is None
    )


def test_ingest_attachment_noops_without_engine():
    # No injected store + no reachable engine -> clean no-op.
    att = {"id": "x", "contentBytes": base64.b64encode(b"y").decode()}
    assert ingest_attachment(att) is None


def test_ingest_attachment_empty_is_noop():
    assert ingest_attachment({}, media_store=_FakeMediaStore()) is None
