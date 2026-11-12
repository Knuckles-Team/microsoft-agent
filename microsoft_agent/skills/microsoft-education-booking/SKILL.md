---
name: microsoft-education-booking
skill_type: skill
description: >-
  Education classes, schools, users and assignments, Bookings businesses and
  appointments, and learning providers and course activities. Use when the agent
  must read or schedule in these services. Do NOT use for the directory.
license: MIT
tags: [microsoft, education, bookings, learning]
metadata:
  author: Knuckles-Team
  version: '0.2.0'
---
# Microsoft Education, Bookings and Learning

Specialist workloads of Microsoft Graph.

## When to use
- Read classes, schools and assignments.
- List Bookings businesses; create an appointment.
- Read learning providers and course activities.

## When NOT to use
- Tenant users and groups: use microsoft-identity-directory.

## Safety
- Writes need MICROSOFT_ALLOW_WRITES and destructive tools also need MICROSOFT_ALLOW_DESTRUCTIVE; the server refuses them otherwise.
- Treat all provider content as sensitive: select the fewest fields, bound result sizes, never echo tokens or credentials.

## Tools (13)

R = read-only, W = changes data, D = destructive (needs approval). Parameters in brackets are the typed inputs.

| Tool | Kind | Inputs |
|---|---|---|
| `microsoft_create_booking_appointment` | W | business_id, data, params |
| `microsoft_get_booking_business` | R | business_id, params |
| `microsoft_get_education_class` | R | class_id, params |
| `microsoft_get_education_school` | R | params, school_id |
| `microsoft_get_learning_provider` | R | params, provider_id |
| `microsoft_list_booking_appointments` | R | business_id, params |
| `microsoft_list_booking_businesses` | R | params |
| `microsoft_list_education_assignments` | R | class_id, params |
| `microsoft_list_education_classes` | R | params |
| `microsoft_list_education_schools` | R | params |
| `microsoft_list_education_users` | R | params |
| `microsoft_list_learning_course_activities` | R | params |
| `microsoft_list_learning_providers` | R | params |

## Tasks

Each task is a tool-use contract the platform serves; steps run in this order.

- `education-booking-booking-appointment`: `microsoft_list_booking_appointments` -> `microsoft_create_booking_appointment`
- `education-booking-booking-business`: `microsoft_get_booking_business` -> `microsoft_list_booking_businesses`
- `education-booking-education-class`: `microsoft_get_education_class` -> `microsoft_list_education_classes`
- `education-booking-education-school`: `microsoft_get_education_school` -> `microsoft_list_education_schools`
- `education-booking-learning-provider`: `microsoft_get_learning_provider` -> `microsoft_list_learning_providers`
- `education-booking-read-records`: `microsoft_list_education_assignments` -> `microsoft_list_education_users` -> `microsoft_list_learning_course_activities`
