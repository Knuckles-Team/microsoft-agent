---
name: microsoft-devices
description: "Microsoft 365 Devices — Directory Devices, Intune Managed Devices & Compliance"
---

# Microsoft 365 Devices

Manage directory devices, Intune managed devices, compliance policies, and device configurations.

## Available Tools

| Tool | Description |
|------|-------------|
| `delete_device` | Delete a device |
| `get_device` | Get a specific device |
| `get_managed_device` | Get a specific managed device |
| `list_device_compliance_policies` | List device compliance policies |
| `list_device_configurations` | List device configuration profiles |
| `list_devices` | List devices registered in the directory |
| `list_managed_devices` | List Intune managed devices |
| `retire_managed_device` | Retire a managed device (remove company data) |
| `wipe_managed_device` | Wipe a managed device (factory reset) |

## Required Permissions
- `Device.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
