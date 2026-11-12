---
name: microsoft-endpoint-management
skill_type: skill
description: >-
  Intune managed devices, compliance and configuration policies, detected apps,
  remote actions (sync, lock, reboot, shut down, wipe, Defender scan) and the
  Windows companion device health and action queue. Use when the agent must
  inspect or act on managed devices. Do NOT use for directory device objects
  (use microsoft-identity-directory).
license: MIT
tags: [microsoft, intune, windows, devices]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Intune and Windows endpoint management

Remote actions affect real devices and are destructive-tier: they need allowlisted devices, time-bound confirmation evidence and approval.

## When to use
- Inspect managed devices, compliance and configuration.
- Sync, lock, scan, reboot, shut down, retire or wipe an allowlisted device.
- Submit a typed Windows companion action and read its result.

## When NOT to use
- Directory device objects: use microsoft-identity-directory.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (19)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `get_intune_configuration` | R | none |
| `get_intune_managed_device` | R | device_id |
| `get_windows_action_result` | R | action_id, device_alias |
| `get_windows_companion_configuration` | R | none |
| `get_windows_device_health` | R | device_alias |
| `list_intune_detected_apps` | R | none |
| `list_intune_managed_devices` | R | none |
| `microsoft_get_managed_device` | R | device_id, params |
| `microsoft_list_device_compliance_policies` | R | params |
| `microsoft_list_device_configurations` | R | params |
| `microsoft_list_managed_devices` | R | params |
| `microsoft_retire_managed_device` | W | device_id, params |
| `microsoft_wipe_managed_device` | W | device_id, params |
| `reboot_intune_device` | D | confirmation, device_id |
| `remote_lock_intune_device` | W | confirmation, device_id |
| `scan_intune_device_with_defender` | W | confirmation, device_id, quick_scan |
| `shut_down_intune_device` | D | confirmation, device_id |
| `submit_windows_action` | W | companion_action, confirmation, device_alias, idempotency_key |
| `sync_intune_device` | W | confirmation, device_id |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `endpoints-intune-device`: `scan_intune_device_with_defender` -> `sync_intune_device` -> `reboot_intune_device`
- `endpoints-intune-managed`: `get_intune_managed_device` -> `list_intune_managed_devices`
- `endpoints-managed-device`: `microsoft_get_managed_device` -> `microsoft_list_managed_devices` -> `microsoft_retire_managed_device` -> `microsoft_wipe_managed_device`
- `endpoints-window-action`: `get_windows_action_result` -> `submit_windows_action`
- `endpoints-read-records`: `get_intune_configuration` -> `get_windows_companion_configuration` -> `get_windows_device_health` -> `list_intune_detected_apps` -> `microsoft_list_device_compliance_policies` -> `microsoft_list_device_configurations`
- `endpoints-change-records`: `remote_lock_intune_device` -> `shut_down_intune_device`
