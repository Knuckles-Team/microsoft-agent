"""Native epistemic-graph ingestion for Microsoft 365 / Graph records.

CONCEPT:AU-KG.ingest.enterprise-source-extractor. The connector natively pushes its
Microsoft Graph data into the ONE epistemic-graph knowledge graph in every modality
that applies (the "maximum ingestion" bar):

* **typed nodes** — mail/chat messages, calendar events, drive files and directory
  users → OWL ``:Message`` / ``:Event`` / ``:File`` / ``:Person`` nodes + links
  (``ingest_messages`` / ``ingest_events`` / ``ingest_files`` / ``ingest_users``).
* **documents** — message and event bodies worth semantic search → ``:Document``
  nodes carrying the text (``ingest_documents``); hub-side enrichment chunks/embeds.
* **blobs** — raw mail attachment bytes → ``:Blob`` + ``:MediaAsset`` via
  :func:`media_store` (see :mod:`microsoft_agent.kg_media`).

This is a **thin mapper** over the shared primitive
``agent_utilities.knowledge_graph.memory.native_ingest``. That primitive is imported
guarded; when it (or the KG stack / a reachable engine) is absent, every entry point
**no-ops** (returns ``None``) using a self-contained txn fallback, so the connector
runs with zero KG infrastructure. Node ids follow ``microsoft:<class>:<externalId>``;
``type`` on each entity matches a class the package's ``ontology_providers`` ttl
federates.
"""

from __future__ import annotations

import logging
import time
from typing import Any

logger = logging.getLogger("microsoft_agent.kg")

_SOURCE = "microsoft-agent"
_DOMAIN = "microsoft"
_DEFAULT_GRAPH = "__commons__"

# --- shared primitive (guarded) --------------------------------------------------
try:  # pragma: no cover - exercised only when the KG stack is installed
    from agent_utilities.knowledge_graph.memory.native_ingest import (
        ingest_documents as _shared_ingest_documents,
    )
    from agent_utilities.knowledge_graph.memory.native_ingest import (
        ingest_entities as _shared_ingest_entities,
    )
    from agent_utilities.knowledge_graph.memory.native_ingest import (
        media_store as _shared_media_store,
    )
except Exception:  # noqa: BLE001 - primitive not yet in installed agent_utilities
    _shared_ingest_entities = None
    _shared_ingest_documents = None
    _shared_media_store = None


# --- self-contained txn fallback -------------------------------------------------
def _client() -> tuple[Any | None, str]:
    """Return ``(engine_client, graph_name)`` or ``(None, "")`` when unavailable."""
    try:
        from agent_utilities.knowledge_graph.core.graph_compute import (
            GraphComputeEngine,
        )
    except Exception as e:  # noqa: BLE001 - KG stack absent
        logger.debug("KG ingest unavailable (import): %s", e)
        return None, ""
    try:
        engine = GraphComputeEngine()
        client = getattr(engine, "_client", None)
        if client is None:
            return None, ""
        return client, (getattr(engine, "graph_name", None) or _DEFAULT_GRAPH)
    except Exception as e:  # noqa: BLE001 - engine unreachable
        logger.debug("KG ingest: engine unreachable: %s", e)
        return None, ""


def _write_nodes(
    client: Any,
    graph: str,
    nodes: list[dict[str, Any]],
    relationships: list[dict[str, Any]] | None,
) -> dict[str, int] | None:
    """Stamp provenance, MERGE the nodes in one txn, then add the edges."""
    nodes = [n for n in nodes if n.get("id")]
    if not nodes:
        return None
    try:
        txn = client.txn.begin(graph=graph)
        for node in nodes:
            props = {k: v for k, v in node.items() if k != "id" and v is not None}
            props.setdefault("source", _SOURCE)
            props.setdefault("domain", _DOMAIN)
            client.txn.add_node(txn, node["id"], props)
        committed = client.txn.commit(txn)
    except Exception as e:  # noqa: BLE001 - engine/txn failure is non-fatal
        logger.warning("KG ingest: txn failed: %s", e)
        return None
    if not committed:
        logger.warning("KG ingest: txn not committed (conflict)")
        return None

    edges = 0
    for rel in relationships or []:
        try:
            client.edges.add(
                rel["source"], rel["target"], {"type": rel.get("type", "RELATED")}
            )
            edges += 1
        except Exception as e:  # noqa: BLE001 - pure edge link, best-effort
            logger.debug("KG ingest: edge skipped: %s", e)

    logger.info("KG ingest: wrote %d nodes, %d edges", len(nodes), edges)
    return {"nodes": len(nodes), "edges": edges}


def ingest_entities(
    entities: list[dict[str, Any]],
    relationships: list[dict[str, Any]] | None = None,
    *,
    source: str = _SOURCE,
    domain: str = _DOMAIN,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Write typed OWL nodes (+ edges). Delegates to the shared primitive when it is
    importable and no explicit ``client`` is injected; otherwise uses the local txn
    fallback (also the path tests drive with a fake client)."""
    entities = [e for e in (entities or []) if e.get("id")]
    if not entities:
        return None
    if client is None and _shared_ingest_entities is not None:
        return _shared_ingest_entities(
            entities, relationships, source=source, domain=domain, graph=graph
        )
    if client is None:
        client, graph = _client()
    if client is None:
        return None
    return _write_nodes(client, graph or _DEFAULT_GRAPH, entities, relationships)


def ingest_documents(
    documents: list[dict[str, Any]],
    *,
    source: str = _SOURCE,
    domain: str = _DOMAIN,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Write text records as ``:Document`` nodes (semantic-search fodder)."""
    documents = [d for d in (documents or []) if d.get("id") and d.get("text")]
    if not documents:
        return None
    if client is None and _shared_ingest_documents is not None:
        return _shared_ingest_documents(
            documents, source=source, domain=domain, graph=graph
        )
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    nodes: list[dict[str, Any]] = []
    for doc in documents:
        node = {k: v for k, v in doc.items() if v is not None}
        node["type"] = "Document"
        node.setdefault("created_at", now)
        nodes.append(node)
    if client is None:
        client, graph = _client()
    if client is None:
        return None
    return _write_nodes(client, graph or _DEFAULT_GRAPH, nodes, None)


def media_store() -> Any | None:
    """Return a live :class:`MediaStore` for raw-blob ingestion, or ``None``."""
    if _shared_media_store is not None:
        try:
            return _shared_media_store()
        except Exception as e:  # noqa: BLE001
            logger.debug("KG ingest: shared media_store failed: %s", e)
    client, _ = _client()
    if client is None:
        return None
    try:
        from agent_utilities.knowledge_graph.core.graph_compute import (
            GraphComputeEngine,
        )
        from agent_utilities.knowledge_graph.memory.media_store import MediaStore

        return MediaStore(GraphComputeEngine())
    except Exception as e:  # noqa: BLE001
        logger.debug("KG ingest: media_store unavailable: %s", e)
        return None


# --- Microsoft Graph record helpers ---------------------------------------------
def _records(resp: Any) -> list[dict[str, Any]]:
    """Normalize a Graph response (``{"value": [...]}`` / list / single) to a list."""
    if resp is None:
        return []
    data = resp
    if isinstance(resp, dict):
        if resp.get("error"):
            return []
        if "value" in resp:
            data = resp["value"]
    if isinstance(data, dict):
        data = [data]
    return [r for r in (data or []) if isinstance(r, dict)]


def _person_id(user: dict[str, Any] | None) -> str | None:
    if not user:
        return None
    uid = user.get("id") or user.get("address") or user.get("userPrincipalName")
    return f"microsoft:Person:{uid}" if uid else None


def _person_node(user: dict[str, Any] | None) -> dict[str, Any] | None:
    pid = _person_id(user)
    if not pid:
        return None
    return {
        "id": pid,
        "type": "Person",
        "name": user.get("displayName") or user.get("name"),
        "emailAddress": user.get("mail")
        or user.get("address")
        or user.get("userPrincipalName"),
        "userPrincipalName": user.get("userPrincipalName"),
        "externalToolId": user.get("id"),
    }


# --- typed-node mappers ----------------------------------------------------------
def ingest_messages(
    messages: list[dict[str, Any]],
    *,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Map Outlook mail / Teams chat message records → ``:Message`` (+ ``:Person``
    sender/recipients + ``:Document`` body) and ingest."""
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, Any]] = []
    documents: list[dict[str, Any]] = []
    seen: set[str] = set()

    for msg in messages or []:
        mid = msg.get("id")
        if not mid:
            continue
        node_id = f"microsoft:Message:{mid}"
        body = msg.get("body") or {}
        body_text = body.get("content") if isinstance(body, dict) else None
        entities.append(
            {
                "id": node_id,
                "type": "Message",
                "subject": msg.get("subject"),
                "text": msg.get("bodyPreview"),
                "webLink": msg.get("webLink"),
                "receivedDateTime": msg.get("receivedDateTime")
                or msg.get("createdDateTime"),
                "externalToolId": mid,
            }
        )
        # sender: mail uses from.emailAddress, chat uses from.user
        sender = msg.get("from") or msg.get("sender") or {}
        if isinstance(sender, dict):
            sender = sender.get("emailAddress") or sender.get("user") or sender
        sender_node = _person_node(sender if isinstance(sender, dict) else None)
        if sender_node:
            if sender_node["id"] not in seen:
                entities.append(sender_node)
                seen.add(sender_node["id"])
            relationships.append(
                {"source": node_id, "target": sender_node["id"], "type": "sentBy"}
            )
        for rcpt in msg.get("toRecipients") or []:
            addr = rcpt.get("emailAddress") if isinstance(rcpt, dict) else None
            rn = _person_node(addr)
            if rn:
                if rn["id"] not in seen:
                    entities.append(rn)
                    seen.add(rn["id"])
                relationships.append(
                    {"source": node_id, "target": rn["id"], "type": "sentTo"}
                )
        if body_text:
            doc_id = f"microsoft:Document:message:{mid}"
            documents.append(
                {
                    "id": doc_id,
                    "text": body_text,
                    "title": msg.get("subject"),
                    "source_uri": msg.get("webLink"),
                }
            )
            relationships.append(
                {"source": node_id, "target": doc_id, "type": "hasBody"}
            )

    res = ingest_entities(entities, relationships, client=client, graph=graph)
    if documents:
        ingest_documents(documents, client=client, graph=graph)
    return res


def ingest_events(
    events: list[dict[str, Any]],
    *,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Map calendar event records → ``:Event`` (+ organizer/attendee ``:Person``)."""
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, Any]] = []
    seen: set[str] = set()

    for ev in events or []:
        eid = ev.get("id")
        if not eid:
            continue
        node_id = f"microsoft:Event:{eid}"
        start = ev.get("start") or {}
        end = ev.get("end") or {}
        entities.append(
            {
                "id": node_id,
                "type": "Event",
                "subject": ev.get("subject"),
                "text": ev.get("bodyPreview"),
                "webLink": ev.get("webLink"),
                "startDateTime": start.get("dateTime")
                if isinstance(start, dict)
                else None,
                "endDateTime": end.get("dateTime") if isinstance(end, dict) else None,
                "externalToolId": eid,
            }
        )
        org = (ev.get("organizer") or {}).get("emailAddress")
        on = _person_node(org)
        if on:
            if on["id"] not in seen:
                entities.append(on)
                seen.add(on["id"])
            relationships.append(
                {"source": node_id, "target": on["id"], "type": "organizedBy"}
            )
        for att in ev.get("attendees") or []:
            addr = att.get("emailAddress") if isinstance(att, dict) else None
            an = _person_node(addr)
            if an:
                if an["id"] not in seen:
                    entities.append(an)
                    seen.add(an["id"])
                relationships.append(
                    {"source": node_id, "target": an["id"], "type": "hasAttendee"}
                )

    return ingest_entities(entities, relationships, client=client, graph=graph)


def ingest_files(
    files: list[dict[str, Any]],
    *,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Map OneDrive/SharePoint drive items → ``:File`` (+ owner ``:Person``)."""
    entities: list[dict[str, Any]] = []
    relationships: list[dict[str, Any]] = []
    seen: set[str] = set()

    for f in files or []:
        fid = f.get("id")
        if not fid:
            continue
        node_id = f"microsoft:File:{fid}"
        file_facet = f.get("file") or {}
        entities.append(
            {
                "id": node_id,
                "type": "File",
                "name": f.get("name"),
                "webLink": f.get("webUrl"),
                "mimeType": file_facet.get("mimeType")
                if isinstance(file_facet, dict)
                else None,
                "sizeBytes": f.get("size"),
                "receivedDateTime": f.get("lastModifiedDateTime"),
                "externalToolId": fid,
            }
        )
        owner = (f.get("createdBy") or {}).get("user")
        on = _person_node(owner)
        if on:
            if on["id"] not in seen:
                entities.append(on)
                seen.add(on["id"])
            relationships.append(
                {"source": node_id, "target": on["id"], "type": "ownedByPerson"}
            )

    return ingest_entities(entities, relationships, client=client, graph=graph)


def ingest_users(
    users: list[dict[str, Any]],
    *,
    client: Any | None = None,
    graph: str | None = None,
) -> dict[str, int] | None:
    """Map directory user / Outlook contact records → ``:Person`` nodes."""
    entities: list[dict[str, Any]] = []
    for u in users or []:
        node = _person_node(u)
        if not node:
            # contacts expose emailAddresses[]; synthesize an address dict
            addrs = u.get("emailAddresses") or []
            addr = addrs[0].get("address") if addrs else None
            node = _person_node(
                {
                    "id": u.get("id"),
                    "displayName": u.get("displayName"),
                    "address": addr,
                }
            )
        if not node:
            continue
        node["jobTitle"] = u.get("jobTitle")
        node["department"] = u.get("department")
        node["companyName"] = u.get("companyName")
        entities.append({k: v for k, v in node.items() if v is not None})
    return ingest_entities(entities, None, client=client, graph=graph)
