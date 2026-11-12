---
name: microsoft-graph-integration
skill_type: skill
description: >-
  Graph change-notification subscriptions, Microsoft Search external
  connections, search queries, and the bounded knowledge-graph ingestion
  projection. Use when the agent must subscribe to changes, manage an external
  connection, search across Microsoft 365, or project structure into the
  knowledge graph. Do NOT use for reading mail or files directly.
license: MIT
tags: [microsoft, subscriptions, search, ingestion]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Graph subscriptions, connections and projection

Integration operations. The ingestion projection returns keyed opaque nodes only; no raw identifiers or content leave the tenant.

## When to use
- Create, renew or delete a change-notification subscription.
- Manage external search connections.
- Run a Microsoft Search query.
- List the ingestion projection.

## When NOT to use
- Reading item content: use the domain skills.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES; a destructive tool also needs a live confirmation (the gateway's approval or the `confirm` argument), and the server refuses it otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (11)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `list_microsoft_ingestion_projection` | W | drive_id, drive_item_id, kind, params_json |
| `microsoft_create_external_connection` | W | data, params |
| `microsoft_create_subscription` | W | data, params |
| `microsoft_delete_external_connection` | D | confirm, connection_id, params |
| `microsoft_delete_subscription` | D | confirm, params, subscription_id |
| `microsoft_get_external_connection` | R | connection_id, params |
| `microsoft_get_subscription` | R | params, subscription_id |
| `microsoft_list_external_connections` | R | params |
| `microsoft_list_subscriptions` | R | params |
| `microsoft_search_query` | R | data, params |
| `microsoft_update_subscription` | W | data, params, subscription_id |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `integration-external-connection`: `microsoft_get_external_connection` -> `microsoft_list_external_connections` -> `microsoft_create_external_connection` -> `microsoft_delete_external_connection`
- `integration-subscription`: `microsoft_get_subscription` -> `microsoft_list_subscriptions` -> `microsoft_create_subscription` -> `microsoft_update_subscription` -> `microsoft_delete_subscription`
- `integration-read-records`: `microsoft_search_query`
- `integration-change-records`: `list_microsoft_ingestion_projection`
