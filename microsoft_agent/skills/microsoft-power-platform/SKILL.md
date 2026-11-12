---
name: microsoft-power-platform
skill_type: skill
description: >-
  Power Automate cloud flows and desktop flows: list, inspect, activate,
  deactivate and run flows, and read desktop flow runs and outputs. Use when the
  agent must operate an automation. Do NOT use for Intune actions (use
  microsoft-endpoint-management).
license: MIT
tags: [microsoft, power-automate, automation]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Power Automate

Running a flow executes real automation: confirm inputs and approval first.

## When to use
- List and inspect cloud and desktop flows.
- Activate, deactivate or run a flow.
- Read a desktop flow run, its outputs, or cancel it.

## When NOT to use
- Device actions: use microsoft-endpoint-management.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (12)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `activate_power_automate_flow` | W | etag, idempotency_key, workflow_id |
| `cancel_power_automate_desktop_flow_run` | D | flow_session_id |
| `deactivate_power_automate_flow` | W | etag, idempotency_key, workflow_id |
| `get_power_automate_desktop_flow_outputs` | R | flow_session_id |
| `get_power_automate_desktop_flow_run` | R | flow_session_id |
| `get_power_automate_desktop_flow_schema` | R | flow_name, kind |
| `get_power_automate_flow` | R | workflow_id |
| `get_power_platform_configuration` | R | none |
| `list_power_automate_desktop_flows` | R | fetch_all, include_unpublished, top |
| `list_power_automate_flows` | R | active_only, fetch_all, top |
| `run_power_automate_desktop_flow` | W | flow_name, inputs, priority, run_mode, timeout_seconds |
| `run_power_automate_flow` | W | flow_name, idempotency_key, payload, timeout_seconds |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `power-platform-power-automate`: `get_power_automate_desktop_flow_outputs` -> `get_power_automate_desktop_flow_run` -> `get_power_automate_desktop_flow_schema` -> `get_power_automate_flow` -> `list_power_automate_desktop_flows` -> `list_power_automate_flows` -> `activate_power_automate_flow` -> `deactivate_power_automate_flow` -> `run_power_automate_desktop_flow` -> `run_power_automate_flow` -> `cancel_power_automate_desktop_flow_run`
- `power-platform-read-records`: `get_power_platform_configuration`
