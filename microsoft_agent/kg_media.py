"""Native epistemic-graph blob ingestion for Microsoft 365 mail attachments.

CONCEPT:AU-KG.ingest.list-durable-media. When a live epistemic-graph engine is
reachable, a mail attachment's raw bytes are stored as a content-addressed **blob**
with a ``:MediaAsset`` graph node (carrying the attachment metadata) via the shared
``MediaStore`` — so the bytes, not just a Graph attachment id, become durable, deduped
and queryable inside the knowledge graph.

Entirely best-effort and engine-guarded: if the KG stack / a live engine is absent
(or :func:`microsoft_agent.kg_ingest.media_store` returns ``None``), every entry point
**no-ops** (returns ``None``), so the connector runs with zero KG infrastructure.
"""

from __future__ import annotations

import base64
import logging
from typing import Any

logger = logging.getLogger("microsoft_agent.kg")


def _media_type(mime: str) -> str:
    if mime.startswith("audio"):
        return "audio"
    if mime.startswith("video"):
        return "video"
    if mime.startswith("image"):
        return "image"
    return "file"


def ingest_attachment(
    attachment: dict[str, Any],
    *,
    message_id: str = "",
    source: str = "microsoft-agent",
    media_store: Any | None = None,
) -> dict[str, Any] | None:
    """Store one Microsoft Graph ``fileAttachment`` as a blob + ``:MediaAsset``.

    ``attachment`` is a Graph attachment record with a base64 ``contentBytes`` field
    (``@odata.type == #microsoft.graph.fileAttachment``). Returns
    ``{asset_id, digest, size_bytes, media_type}`` on success, or ``None`` when there
    is no engine, no bytes, or the store failed (never raises). ``media_store`` may be
    injected (tests); otherwise one is resolved via :func:`kg_ingest.media_store`.
    """
    if not attachment:
        return None
    raw = attachment.get("contentBytes")
    if not raw:
        return None
    try:
        data = base64.b64decode(raw) if isinstance(raw, str) else bytes(raw)
    except Exception as e:  # noqa: BLE001 - malformed payload
        logger.warning("KG media ingest: bad contentBytes: %s", e)
        return None
    if not data:
        return None

    store = media_store
    if store is None:
        from microsoft_agent.kg_ingest import media_store as _resolve

        store = _resolve()
    if store is None:
        return None

    mime = attachment.get("contentType") or "application/octet-stream"
    name = attachment.get("name") or attachment.get("id") or "attachment"
    extra = {
        "attachment_id": attachment.get("id"),
        "message_id": message_id,
        "content_type": mime,
    }
    extra = {k: v for k, v in extra.items() if v is not None}

    try:
        stored = store.store_media(
            data,
            media_type=_media_type(mime),
            mime_type=mime,
            source=source,
            name=name,
            extra=extra,
        )
    except Exception as e:  # noqa: BLE001 - engine/store failure is non-fatal
        logger.warning("KG media ingest: store_media failed: %s", e)
        return None
    if stored is None:
        return None

    logger.info(
        "KG media ingest: stored %s (%d bytes) as asset %s",
        name,
        len(data),
        getattr(stored, "asset_id", "?"),
    )
    return {
        "asset_id": getattr(stored, "asset_id", None),
        "digest": getattr(stored, "digest", None),
        "size_bytes": len(data),
        "media_type": _media_type(mime),
    }
