---
name: microsoft-solutions
description: "Microsoft 365 Solutions — Booking Businesses, Appointments & Virtual Events"
---

# Microsoft 365 Solutions

Manage booking businesses, appointments, and virtual events.

## Available Tools

| Tool | Description |
|------|-------------|
| `create_booking_appointment` | Create a booking appointment |
| `get_booking_business` | Get a specific booking business |
| `list_booking_appointments` | List appointments for a booking business |
| `list_booking_businesses` | List booking businesses |
| `list_virtual_events` | List virtual event townhalls |

## Required Permissions
- `Bookings.ReadWrite.All`

## Error Handling
All tools return `{"error": "<message>"}` on failure.
