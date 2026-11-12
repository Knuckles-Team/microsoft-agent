"""Surface contracts: tool use, links, data locations and business roles.

Generated from the served tool surface and edited as data; every tool is called by a step.
"""

# ruff: noqa: E501
from scion_sdk.surface_contracts import (
    BusinessRole,
    DataLocation,
    Refusal,
    SurfaceLink,
    ToolStep,
    ToolUseContract,
)

CONTRACTS = (
    ToolUseContract(
        task="calendar-event",
        description="Read, create or change and remove event through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_calendar_event",
                params=("event_id", "params", "timezone"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_calendar_events",
                params=("params", "timezone"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_calendar_event",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_calendar_event",
                params=("data", "event_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_calendar_event",
                params=("confirm", "event_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="calendar-online-meeting",
        description="Read, create or change and remove online meeting through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_online_meeting",
                params=("meeting_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_online_meetings",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_online_meeting",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_online_meeting",
                params=("data", "meeting_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_online_meeting",
                params=("confirm", "meeting_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="calendar-place",
        description="Read and create or change place through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_place",
                params=("params", "place_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_place",
                params=("data", "params", "place_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="calendar-specific-calendar",
        description="Read, create or change and remove specific calendar through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_specific_calendar_event",
                params=("calendar_id", "event_id", "params", "timezone"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_specific_calendar_events",
                params=("calendar_id", "params", "timezone"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_specific_calendar_event",
                params=("calendar_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_update_specific_calendar_event",
                params=("calendar_id", "data", "event_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_specific_calendar_event",
                params=("calendar_id", "confirm", "event_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="contacts-outlook-contact",
        description="Read, create or change and remove outlook contact through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_outlook_contact",
                params=("contact_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_outlook_contacts",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_outlook_contact",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_outlook_contact",
                params=("contact_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_outlook_contact",
                params=("confirm", "contact_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="documents-powerpoint-slide",
        description="Read, create or change and remove powerpoint slide through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="list_powerpoint_slides_in_office",
                params=("session_id", "wait_timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="add_powerpoint_slide_in_office",
                params=("session_id", "wait_timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="delete_powerpoint_slide_in_office",
                params=("session_id", "slide_number", "wait_timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="documents-word-selection",
        description="Read and create or change word selection through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_word_selection_from_office",
                params=("session_id", "wait_timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="write_word_selection_in_office",
                params=("content", "mode", "session_id", "wait_timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-booking-appointment",
        description="Read and create or change booking booking appointment through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_booking_appointments",
                params=("business_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_booking_appointment",
                params=("business_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-booking-business",
        description="Read booking booking business through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_booking_business",
                params=("business_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_booking_businesses",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-education-class",
        description="Read booking education class through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_education_class",
                params=("class_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_education_classes",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-education-school",
        description="Read booking education school through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_education_school",
                params=("params", "school_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_education_schools",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-learning-provider",
        description="Read booking learning provider through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_learning_provider",
                params=("params", "provider_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_learning_providers",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-intune-device",
        description="Create or change and remove intune device through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="scan_intune_device_with_defender",
                params=("confirmation", "device_id", "quick_scan"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="sync_intune_device",
                params=("confirmation", "device_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="reboot_intune_device",
                params=("confirmation", "device_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-intune-managed",
        description="Read intune managed through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_intune_managed_device",
                params=("device_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="list_intune_managed_devices",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-managed-device",
        description="Read and create or change managed device through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_managed_device",
                params=("device_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_managed_devices",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_retire_managed_device",
                params=("device_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_wipe_managed_device",
                params=("device_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-window-action",
        description="Read and create or change window action through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_windows_action_result",
                params=("action_id", "device_alias"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="submit_windows_action",
                params=(
                    "companion_action",
                    "confirmation",
                    "device_alias",
                    "idempotency_key",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-admin-sharepoint",
        description="Read and create or change admin sharepoint through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_admin_sharepoint",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_admin_sharepoint",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-file-storage",
        description="Read and create or change file storage through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_file_storage_container",
                params=("container_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_file_storage_containers",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_file_storage_container",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-onedrive-file",
        description="Create or change and remove onedrive file through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_download_onedrive_file_content",
                params=("driveItem_id", "drive_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_onedrive_file",
                params=("confirm", "driveItem_id", "drive_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-sharepoint-site",
        description="Read sharepoint site through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_sharepoint_site_by_path",
                params=("params", "path", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_sharepoint_site_list_item",
                params=("listItem_id", "list_id", "params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_sharepoint_sites_delta",
                params=("continuation_url", "fetch_all", "max_pages", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_sharepoint_site_list_items",
                params=("list_id", "params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-site",
        description="Read site through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_site",
                params=("params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_sites",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-site-drive",
        description="Read site drive through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_site_drive_by_id",
                params=("drive_id", "params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_site_drives",
                params=("params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-site-item",
        description="Read site item through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_site_item",
                params=("baseItem_id", "params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_site_items",
                params=("params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-site-list",
        description="Read site list through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_site_list",
                params=("list_id", "params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_site_lists",
                params=("params", "site_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-access-review",
        description="Read access review through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_access_review",
                params=("params", "review_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_access_reviews",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-agreement",
        description="Read, create or change and remove agreement through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_agreement",
                params=("agreement_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_agreements",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_agreement",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_agreement",
                params=("agreement_id", "confirm", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-application",
        description="Read, create or change and remove application through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_application",
                params=("app_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_applications",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_application",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_application",
                params=("app_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_application",
                params=("app_id", "confirm", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-application-password",
        description="Create or change and remove application password through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_add_application_password",
                params=("app_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_remove_application_password",
                params=("app_id", "confirm", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-conditional-access",
        description="Read, create or change and remove conditional access through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_conditional_access_policy",
                params=("params", "policy_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_conditional_access_policies",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_conditional_access_policy",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_conditional_access_policy",
                params=("data", "params", "policy_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_conditional_access_policy",
                params=("confirm", "params", "policy_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-delegated-admin",
        description="Read delegated admin through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_delegated_admin_relationship",
                params=("params", "rel_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_delegated_admin_relationships",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-deleted-item",
        description="Read and create or change deleted item through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_deleted_items",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_restore_deleted_item",
                params=("object_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-device",
        description="Read and remove device through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_device",
                params=("device_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_devices",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_device",
                params=("confirm", "device_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-directory-object",
        description="Read directory object through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_directory_object",
                params=("object_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_directory_objects",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-directory-role",
        description="Read directory role through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_directory_role",
                params=("params", "role_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_directory_role_templates",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_directory_roles",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-domain",
        description="Read, create or change and remove domain through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_domain",
                params=("domain_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_domains",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_domain",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_verify_domain",
                params=("domain_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_domain",
                params=("confirm", "domain_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-group",
        description="Read, create or change and remove group through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_group",
                params=("group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_groups",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_group",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_group",
                params=("data", "group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_group",
                params=("confirm", "group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-group-member",
        description="Read, create or change and remove group member through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_group_members",
                params=("group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_add_group_member",
                params=("data", "group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_remove_group_member",
                params=("confirm", "group_id", "member_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-org-branding",
        description="Read and create or change org branding through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_org_branding",
                params=("org_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_org_branding",
                params=("data", "org_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-organization",
        description="Read and create or change organization through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_organization",
                params=("org_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_organization",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_organization",
                params=("data", "org_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-role-assignment",
        description="Read and create or change role assignment through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_role_assignment",
                params=("assignment_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_role_assignments",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_role_assignment",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-role-definition",
        description="Read role definition through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_role_definition",
                params=("params", "role_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_role_definitions",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-service-principal",
        description="Read, create or change and remove service principal through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_service_principal",
                params=("params", "sp_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_service_principals",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_service_principal",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_service_principal",
                params=("data", "params", "sp_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_service_principal",
                params=("confirm", "params", "sp_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="integration-external-connection",
        description="Read, create or change and remove external connection through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_external_connection",
                params=("connection_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_external_connections",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_external_connection",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_external_connection",
                params=("confirm", "connection_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="integration-subscription",
        description="Read, create or change and remove subscription through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_subscription",
                params=("params", "subscription_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_subscriptions",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_subscription",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_subscription",
                params=("data", "params", "subscription_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_subscription",
                params=("confirm", "params", "subscription_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-attachment",
        description="Read, create or change and remove attachment through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_mail_attachment",
                params=("attachment_id", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_mail_attachments",
                params=("message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_add_mail_attachment",
                params=("data", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_mail_attachment",
                params=("attachment_id", "confirm", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-folder",
        description="Read folder through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_mail_folder_messages",
                params=("mailFolder_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_mail_folders",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-message",
        description="Read, create or change and remove message through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_mail_message",
                params=("message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_mail_messages",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_move_mail_message",
                params=("data", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_update_mail_message",
                params=("data", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_mail_message",
                params=("confirm", "message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-shared-mailbox",
        description="Read and create or change shared mailbox through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_shared_mailbox_message",
                params=("message_id", "params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_shared_mailbox_folder_messages",
                params=("mailFolder_id", "params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_shared_mailbox_messages",
                params=("params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_send_shared_mailbox_mail",
                params=("data", "params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-excel-range",
        description="Read and create or change apps excel range through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_excel_range",
                params=("address", "drive_id", "item_id", "params", "worksheet_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_format_excel_range",
                params=(
                    "address",
                    "data",
                    "drive_id",
                    "item_id",
                    "params",
                    "worksheet_id",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_sort_excel_range",
                params=(
                    "address",
                    "data",
                    "drive_id",
                    "item_id",
                    "params",
                    "worksheet_id",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-excel-table",
        description="Read apps excel table through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_excel_table",
                params=("drive_id", "item_id", "params", "table_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_excel_tables",
                params=("drive_id", "item_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-excel-worksheet",
        description="Read apps excel worksheet through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_excel_worksheet",
                params=("drive_id", "item_id", "params", "worksheet_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_excel_worksheets",
                params=("drive_id", "item_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-onenote-notebook",
        description="Read apps onenote notebook through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_onenote_notebook_sections",
                params=("notebook_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_onenote_notebooks",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-onenote-page",
        description="Read and create or change apps onenote page through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_onenote_page_content",
                params=("onenotePage_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_onenote_page",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-planner-task",
        description="Read and create or change apps planner task through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_planner_task",
                params=("params", "plannerTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_planner_tasks",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_planner_task",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_planner_task",
                params=("data", "etag", "params", "plannerTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_update_planner_task_details",
                params=("data", "etag", "params", "plannerTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-todo-task",
        description="Read, create or change and remove apps todo task through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_todo_task",
                params=("params", "todoTaskList_id", "todoTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_todo_task_lists",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_todo_tasks",
                params=("params", "todoTaskList_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_todo_task",
                params=("data", "params", "todoTaskList_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_update_todo_task",
                params=("data", "params", "todoTaskList_id", "todoTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_delete_todo_task",
                params=("confirm", "params", "todoTaskList_id", "todoTask_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="power-platform-power-automate",
        description="Read, create or change and remove platform power automate through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_power_automate_desktop_flow_outputs",
                params=("flow_session_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="get_power_automate_desktop_flow_run",
                params=("flow_session_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="get_power_automate_desktop_flow_schema",
                params=("flow_name", "kind"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="get_power_automate_flow",
                params=("workflow_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="list_power_automate_desktop_flows",
                params=("fetch_all", "include_unpublished", "top"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="list_power_automate_flows",
                params=("active_only", "fetch_all", "top"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="activate_power_automate_flow",
                params=("etag", "idempotency_key", "workflow_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="deactivate_power_automate_flow",
                params=("etag", "idempotency_key", "workflow_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="run_power_automate_desktop_flow",
                params=(
                    "flow_name",
                    "inputs",
                    "priority",
                    "run_mode",
                    "timeout_seconds",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="run_power_automate_flow",
                params=("flow_name", "idempotency_key", "payload", "timeout_seconds"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="cancel_power_automate_desktop_flow_run",
                params=("flow_session_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="print-document",
        description="Read and create or change document through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_create_print_document_upload_session",
                params=(
                    "content_type",
                    "document_name",
                    "print_document_id",
                    "print_job_id",
                    "printer_id",
                    "size",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_submit_print_document",
                params=("printer_id", "submission", "upload_transport"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="print-job",
        description="Read and create or change job through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_print_jobs",
                params=("params", "printer_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_print_job",
                params=("data", "params", "printer_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_start_print_job",
                params=("print_job_id", "printer_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="printer",
        description="Read printer through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_printer",
                params=("params", "printer_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_printers",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="reports-service-health",
        description="Read service health through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_service_health",
                params=("params", "service_name"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_service_health_issue",
                params=("issue_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_service_health",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_service_health_issues",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="reports-service-update",
        description="Read service update through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_service_update_message",
                params=("message_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_service_update_messages",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-directory-audit",
        description="Read directory audit through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_directory_audit",
                params=("audit_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_directory_audits",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-risk-detection",
        description="Read risk detection through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_risk_detection",
                params=("params", "risk_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_risk_detections",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-risky-user",
        description="Read and create or change risky user through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_risky_user",
                params=("params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_risky_users",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_dismiss_risky_user",
                params=("params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-alert",
        description="Read and create or change alert through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_security_alert",
                params=("alert_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_security_alerts",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_security_alert",
                params=("alert_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-incident",
        description="Read and create or change incident through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_security_incident",
                params=("incident_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_security_incidents",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_update_security_incident",
                params=("data", "incident_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-sensitivity-label",
        description="Read sensitivity label through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_sensitivity_label",
                params=("label_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_sensitivity_labels",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-sign-log",
        description="Read sign log through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_sign_in_log",
                params=("params", "sign_in_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_sign_in_logs",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-subject-right",
        description="Read and create or change subject right through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_subject_rights_request",
                params=("params", "request_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_subject_rights_requests",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_create_subject_rights_request",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-threat-intelligence",
        description="Read threat intelligence through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_threat_intelligence_host",
                params=("host_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_threat_intelligence_hosts",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="session-login",
        description="Create or change login through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_login",
                params=("force",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_verify_login",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-call-record",
        description="Read call record through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_call_record",
                params=("call_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_call_records",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-channel-message",
        description="Read and create or change channel message through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_channel_message",
                params=("channel_id", "chatMessage_id", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_channel_message_replies",
                params=("channel_id", "chatMessage_id", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_channel_messages",
                params=("channel_id", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_send_channel_message",
                params=("channel_id", "data", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_reply_to_channel_message",
                params=("channel_id", "chatMessage_id", "data", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-chat",
        description="Read chat through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_chat",
                params=("chat_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_chats",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-chat-message",
        description="Read and create or change chat message through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_chat_message",
                params=("chatMessage_id", "chat_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_chat_message_replies",
                params=("chatMessage_id", "chat_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_chat_messages",
                params=("chat_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_send_chat_message",
                params=("chat_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="microsoft_reply_to_chat_message",
                params=("chatMessage_id", "chat_id", "data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-presence",
        description="Read presence through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_my_presence",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_presence",
                params=("params", "user_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_presences",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-team-channel",
        description="Read team channel through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_team_channel",
                params=("channel_id", "params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_team_channels",
                params=("params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="calendar-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_find_meeting_times",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_calendar_view",
                params=("params", "timezone"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_calendars",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_room_lists",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_rooms",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_virtual_events",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="documents-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_document_capabilities",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="get_office_command_result",
                params=("command_id",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="list_office_sessions",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="documents-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="create_office_pairing",
                params=("expected_host", "label"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="generate_and_upload_powerpoint_presentation",
                params=("conflict_behavior", "destination_path", "drive_id", "request"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="generate_and_upload_word_document",
                params=("conflict_behavior", "destination_path", "drive_id", "request"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="generate_powerpoint_presentation",
                params=("request",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="generate_word_document",
                params=("request",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="insert_powerpoint_text_box_in_office",
                params=(
                    "height",
                    "left",
                    "session_id",
                    "slide_number",
                    "text",
                    "top",
                    "wait_timeout_seconds",
                    "width",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="replace_word_placeholders_in_office",
                params=(
                    "placeholder",
                    "replacement",
                    "session_id",
                    "wait_timeout_seconds",
                ),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="education-booking-read-records",
        description="Read booking read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_education_assignments",
                params=("class_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_education_users",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_learning_course_activities",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_intune_configuration",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="get_windows_companion_configuration",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="get_windows_device_health",
                params=("device_alias",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="list_intune_detected_apps",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_device_compliance_policies",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_device_configurations",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="endpoints-change-records",
        description="Create or change and remove change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="remote_lock_intune_device",
                params=("confirmation", "device_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
            ToolStep(
                tool="shut_down_intune_device",
                params=("confirmation", "device_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the destructive call is not confirmed (no gateway approval and no confirm argument)",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_drive_root_item",
                params=("drive_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_onedrive_usage_report",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_sharepoint_activity_report",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_drives",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_folder_files",
                params=("driveItem_id", "drive_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_group_drives",
                params=("group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="files-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_upload_file_content",
                params=("data", "driveItem_id", "drive_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-read-records-1",
        description="Read read records 1 through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_admin_consent_policy",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_authorization_policy",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_me",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_domain_service_configuration_records",
                params=("domain_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_entitlement_access_packages",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_group_conversations",
                params=("group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_group_owners",
                params=("group_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_lifecycle_workflows",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_permission_grant_policies",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_plan_tasks",
                params=("params", "plannerPlan_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_provisioning_logs",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_token_issuance_policies",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-read-records-2",
        description="Read read records 2 through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_token_lifetime_policies",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_users",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="identity-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_create_invitation",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="integration-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_search_query",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="integration-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="list_microsoft_ingestion_projection",
                params=("drive_id", "drive_item_id", "kind", "params_json"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_email_activity_report",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_mailbox_usage_report",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="mail-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_create_draft_email",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_send_mail",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-read-records",
        description="Read apps read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_excel_workbook",
                params=("drive_id", "item_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_planner_plan",
                params=("params", "plannerPlan_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_onenote_section_pages",
                params=("onenoteSection_id", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="office-apps-change-records",
        description="Create or change apps change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_create_excel_chart",
                params=("data", "drive_id", "item_id", "params", "worksheet_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                    Refusal(code="NOT_FOUND", when="the target record does not exist"),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="power-platform-read-records",
        description="Read platform read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="get_power_platform_configuration",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="print-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_print_shares",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="reports-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_office365_active_users",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_secure_scores",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="security-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_run_hunting_query",
                params=("data", "params"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="session-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_list_accounts",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_search_tools",
                params=("limit", "query"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="session-change-records",
        description="Create or change change records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="login_power_platform",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="login_windows_companion",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_logout",
                params=(),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="writes are disabled (MICROSOFT_ALLOW_WRITES is off) or the caller lacks the Graph scope",
                    ),
                ),
            ),
        ),
    ),
    ToolUseContract(
        task="teams-read-records",
        description="Read read records through their typed operations, in dependency order.",
        steps=(
            ToolStep(
                tool="microsoft_get_team",
                params=("params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_get_teams_user_activity",
                params=("params", "period"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_joined_teams",
                params=("params",),
                result=("result",),
                refusals=(
                    Refusal(
                        code="FORBIDDEN",
                        when="the caller lacks the permission this read requires",
                    ),
                ),
            ),
            ToolStep(
                tool="microsoft_list_team_members",
                params=("params", "team_id"),
                result=("result",),
                refusals=(
                    Refusal(
                        code="NOT_FOUND",
                        when="the target record does not exist or is not visible to the caller",
                    ),
                ),
            ),
        ),
    ),
)

LINKS = (
    SurfaceLink(
        "calendar-event",
        tools=(
            "microsoft_get_calendar_event",
            "microsoft_list_calendar_events",
            "microsoft_create_calendar_event",
            "microsoft_update_calendar_event",
            "microsoft_delete_calendar_event",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-calendar-scheduling",),
    ),
    SurfaceLink(
        "calendar-online-meeting",
        tools=(
            "microsoft_get_online_meeting",
            "microsoft_list_online_meetings",
            "microsoft_create_online_meeting",
            "microsoft_update_online_meeting",
            "microsoft_delete_online_meeting",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-calendar-scheduling",),
    ),
    SurfaceLink(
        "calendar-place",
        tools=("microsoft_get_place", "microsoft_update_place"),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-calendar-scheduling",),
    ),
    SurfaceLink(
        "calendar-specific-calendar",
        tools=(
            "microsoft_get_specific_calendar_event",
            "microsoft_list_specific_calendar_events",
            "microsoft_create_specific_calendar_event",
            "microsoft_update_specific_calendar_event",
            "microsoft_delete_specific_calendar_event",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-calendar-scheduling",),
    ),
    SurfaceLink(
        "contacts-outlook-contact",
        tools=(
            "microsoft_get_outlook_contact",
            "microsoft_list_outlook_contacts",
            "microsoft_create_outlook_contact",
            "microsoft_update_outlook_contact",
            "microsoft_delete_outlook_contact",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-contacts",),
    ),
    SurfaceLink(
        "documents-powerpoint-slide",
        tools=(
            "list_powerpoint_slides_in_office",
            "add_powerpoint_slide_in_office",
            "delete_powerpoint_slide_in_office",
        ),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-document-authoring",),
    ),
    SurfaceLink(
        "documents-word-selection",
        tools=("get_word_selection_from_office", "write_word_selection_in_office"),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-document-authoring",),
    ),
    SurfaceLink(
        "education-booking-booking-appointment",
        tools=(
            "microsoft_list_booking_appointments",
            "microsoft_create_booking_appointment",
        ),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "education-booking-booking-business",
        tools=("microsoft_get_booking_business", "microsoft_list_booking_businesses"),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "education-booking-education-class",
        tools=("microsoft_get_education_class", "microsoft_list_education_classes"),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "education-booking-education-school",
        tools=("microsoft_get_education_school", "microsoft_list_education_schools"),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "education-booking-learning-provider",
        tools=("microsoft_get_learning_provider", "microsoft_list_learning_providers"),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "endpoints-intune-device",
        tools=(
            "scan_intune_device_with_defender",
            "sync_intune_device",
            "reboot_intune_device",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "endpoints-intune-managed",
        tools=("get_intune_managed_device", "list_intune_managed_devices"),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "endpoints-managed-device",
        tools=(
            "microsoft_get_managed_device",
            "microsoft_list_managed_devices",
            "microsoft_retire_managed_device",
            "microsoft_wipe_managed_device",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "endpoints-window-action",
        tools=("get_windows_action_result", "submit_windows_action"),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "files-admin-sharepoint",
        tools=("microsoft_get_admin_sharepoint", "microsoft_update_admin_sharepoint"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-file-storage",
        tools=(
            "microsoft_get_file_storage_container",
            "microsoft_list_file_storage_containers",
            "microsoft_create_file_storage_container",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-onedrive-file",
        tools=(
            "microsoft_download_onedrive_file_content",
            "microsoft_delete_onedrive_file",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-sharepoint-site",
        tools=(
            "microsoft_get_sharepoint_site_by_path",
            "microsoft_get_sharepoint_site_list_item",
            "microsoft_get_sharepoint_sites_delta",
            "microsoft_list_sharepoint_site_list_items",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-site",
        tools=("microsoft_get_site", "microsoft_list_sites"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-site-drive",
        tools=("microsoft_get_site_drive_by_id", "microsoft_list_site_drives"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-site-item",
        tools=("microsoft_get_site_item", "microsoft_list_site_items"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-site-list",
        tools=("microsoft_get_site_list", "microsoft_list_site_lists"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "identity-access-review",
        tools=("microsoft_get_access_review", "microsoft_list_access_reviews"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-agreement",
        tools=(
            "microsoft_get_agreement",
            "microsoft_list_agreements",
            "microsoft_create_agreement",
            "microsoft_delete_agreement",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-application",
        tools=(
            "microsoft_get_application",
            "microsoft_list_applications",
            "microsoft_create_application",
            "microsoft_update_application",
            "microsoft_delete_application",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-application-password",
        tools=(
            "microsoft_add_application_password",
            "microsoft_remove_application_password",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-conditional-access",
        tools=(
            "microsoft_get_conditional_access_policy",
            "microsoft_list_conditional_access_policies",
            "microsoft_create_conditional_access_policy",
            "microsoft_update_conditional_access_policy",
            "microsoft_delete_conditional_access_policy",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-delegated-admin",
        tools=(
            "microsoft_get_delegated_admin_relationship",
            "microsoft_list_delegated_admin_relationships",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-deleted-item",
        tools=("microsoft_list_deleted_items", "microsoft_restore_deleted_item"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-device",
        tools=(
            "microsoft_get_device",
            "microsoft_list_devices",
            "microsoft_delete_device",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-directory-object",
        tools=("microsoft_get_directory_object", "microsoft_list_directory_objects"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-directory-role",
        tools=(
            "microsoft_get_directory_role",
            "microsoft_list_directory_role_templates",
            "microsoft_list_directory_roles",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-domain",
        tools=(
            "microsoft_get_domain",
            "microsoft_list_domains",
            "microsoft_create_domain",
            "microsoft_verify_domain",
            "microsoft_delete_domain",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-group",
        tools=(
            "microsoft_get_group",
            "microsoft_list_groups",
            "microsoft_create_group",
            "microsoft_update_group",
            "microsoft_delete_group",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-group-member",
        tools=(
            "microsoft_list_group_members",
            "microsoft_add_group_member",
            "microsoft_remove_group_member",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-org-branding",
        tools=("microsoft_get_org_branding", "microsoft_update_org_branding"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-organization",
        tools=(
            "microsoft_get_organization",
            "microsoft_list_organization",
            "microsoft_update_organization",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-role-assignment",
        tools=(
            "microsoft_get_role_assignment",
            "microsoft_list_role_assignments",
            "microsoft_create_role_assignment",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-role-definition",
        tools=("microsoft_get_role_definition", "microsoft_list_role_definitions"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-service-principal",
        tools=(
            "microsoft_get_service_principal",
            "microsoft_list_service_principals",
            "microsoft_create_service_principal",
            "microsoft_update_service_principal",
            "microsoft_delete_service_principal",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "integration-external-connection",
        tools=(
            "microsoft_get_external_connection",
            "microsoft_list_external_connections",
            "microsoft_create_external_connection",
            "microsoft_delete_external_connection",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-graph-integration",),
    ),
    SurfaceLink(
        "integration-subscription",
        tools=(
            "microsoft_get_subscription",
            "microsoft_list_subscriptions",
            "microsoft_create_subscription",
            "microsoft_update_subscription",
            "microsoft_delete_subscription",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-graph-integration",),
    ),
    SurfaceLink(
        "mail-attachment",
        tools=(
            "microsoft_get_mail_attachment",
            "microsoft_list_mail_attachments",
            "microsoft_add_mail_attachment",
            "microsoft_delete_mail_attachment",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "mail-folder",
        tools=("microsoft_list_mail_folder_messages", "microsoft_list_mail_folders"),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "mail-message",
        tools=(
            "microsoft_get_mail_message",
            "microsoft_list_mail_messages",
            "microsoft_move_mail_message",
            "microsoft_update_mail_message",
            "microsoft_delete_mail_message",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "mail-shared-mailbox",
        tools=(
            "microsoft_get_shared_mailbox_message",
            "microsoft_list_shared_mailbox_folder_messages",
            "microsoft_list_shared_mailbox_messages",
            "microsoft_send_shared_mailbox_mail",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "office-apps-excel-range",
        tools=(
            "microsoft_get_excel_range",
            "microsoft_format_excel_range",
            "microsoft_sort_excel_range",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-excel-table",
        tools=("microsoft_get_excel_table", "microsoft_list_excel_tables"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-excel-worksheet",
        tools=("microsoft_get_excel_worksheet", "microsoft_list_excel_worksheets"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-onenote-notebook",
        tools=(
            "microsoft_list_onenote_notebook_sections",
            "microsoft_list_onenote_notebooks",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-onenote-page",
        tools=("microsoft_get_onenote_page_content", "microsoft_create_onenote_page"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-planner-task",
        tools=(
            "microsoft_get_planner_task",
            "microsoft_list_planner_tasks",
            "microsoft_create_planner_task",
            "microsoft_update_planner_task",
            "microsoft_update_planner_task_details",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-todo-task",
        tools=(
            "microsoft_get_todo_task",
            "microsoft_list_todo_task_lists",
            "microsoft_list_todo_tasks",
            "microsoft_create_todo_task",
            "microsoft_update_todo_task",
            "microsoft_delete_todo_task",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "power-platform-power-automate",
        tools=(
            "get_power_automate_desktop_flow_outputs",
            "get_power_automate_desktop_flow_run",
            "get_power_automate_desktop_flow_schema",
            "get_power_automate_flow",
            "list_power_automate_desktop_flows",
            "list_power_automate_flows",
            "activate_power_automate_flow",
            "deactivate_power_automate_flow",
            "run_power_automate_desktop_flow",
            "run_power_automate_flow",
            "cancel_power_automate_desktop_flow_run",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-power-platform",),
    ),
    SurfaceLink(
        "print-document",
        tools=(
            "microsoft_create_print_document_upload_session",
            "microsoft_submit_print_document",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-universal-print",),
    ),
    SurfaceLink(
        "print-job",
        tools=(
            "microsoft_list_print_jobs",
            "microsoft_create_print_job",
            "microsoft_start_print_job",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-universal-print",),
    ),
    SurfaceLink(
        "printer",
        tools=("microsoft_get_printer", "microsoft_list_printers"),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-universal-print",),
    ),
    SurfaceLink(
        "reports-service-health",
        tools=(
            "microsoft_get_service_health",
            "microsoft_get_service_health_issue",
            "microsoft_list_service_health",
            "microsoft_list_service_health_issues",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-service-reports",),
    ),
    SurfaceLink(
        "reports-service-update",
        tools=(
            "microsoft_get_service_update_message",
            "microsoft_list_service_update_messages",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-service-reports",),
    ),
    SurfaceLink(
        "security-directory-audit",
        tools=("microsoft_get_directory_audit", "microsoft_list_directory_audits"),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-risk-detection",
        tools=("microsoft_get_risk_detection", "microsoft_list_risk_detections"),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-risky-user",
        tools=(
            "microsoft_get_risky_user",
            "microsoft_list_risky_users",
            "microsoft_dismiss_risky_user",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-alert",
        tools=(
            "microsoft_get_security_alert",
            "microsoft_list_security_alerts",
            "microsoft_update_security_alert",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-incident",
        tools=(
            "microsoft_get_security_incident",
            "microsoft_list_security_incidents",
            "microsoft_update_security_incident",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-sensitivity-label",
        tools=("microsoft_get_sensitivity_label", "microsoft_list_sensitivity_labels"),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-sign-log",
        tools=("microsoft_get_sign_in_log", "microsoft_list_sign_in_logs"),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-subject-right",
        tools=(
            "microsoft_get_subject_rights_request",
            "microsoft_list_subject_rights_requests",
            "microsoft_create_subject_rights_request",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-threat-intelligence",
        tools=(
            "microsoft_get_threat_intelligence_host",
            "microsoft_list_threat_intelligence_hosts",
        ),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "session-login",
        tools=("microsoft_login", "microsoft_verify_login"),
        prompts=(
            "microsoft-agent-operator",
            "microsoft-mailbox-coordinator",
            "microsoft-collaboration-manager",
            "microsoft-identity-administrator",
            "microsoft-security-analyst",
            "microsoft-endpoint-engineer",
            "microsoft-document-author",
        ),
        skills=("microsoft-session-management", "microsoft-agent-operations"),
    ),
    SurfaceLink(
        "teams-call-record",
        tools=("microsoft_get_call_record", "microsoft_list_call_records"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "teams-channel-message",
        tools=(
            "microsoft_get_channel_message",
            "microsoft_list_channel_message_replies",
            "microsoft_list_channel_messages",
            "microsoft_send_channel_message",
            "microsoft_reply_to_channel_message",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "teams-chat",
        tools=("microsoft_get_chat", "microsoft_list_chats"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "teams-chat-message",
        tools=(
            "microsoft_get_chat_message",
            "microsoft_list_chat_message_replies",
            "microsoft_list_chat_messages",
            "microsoft_send_chat_message",
            "microsoft_reply_to_chat_message",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "teams-presence",
        tools=(
            "microsoft_get_my_presence",
            "microsoft_get_presence",
            "microsoft_list_presences",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "teams-team-channel",
        tools=("microsoft_get_team_channel", "microsoft_list_team_channels"),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
    SurfaceLink(
        "calendar-read-records",
        tools=(
            "microsoft_find_meeting_times",
            "microsoft_get_calendar_view",
            "microsoft_list_calendars",
            "microsoft_list_room_lists",
            "microsoft_list_rooms",
            "microsoft_list_virtual_events",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-calendar-scheduling",),
    ),
    SurfaceLink(
        "documents-read-records",
        tools=(
            "get_document_capabilities",
            "get_office_command_result",
            "list_office_sessions",
        ),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-document-authoring",),
    ),
    SurfaceLink(
        "documents-change-records",
        tools=(
            "create_office_pairing",
            "generate_and_upload_powerpoint_presentation",
            "generate_and_upload_word_document",
            "generate_powerpoint_presentation",
            "generate_word_document",
            "insert_powerpoint_text_box_in_office",
            "replace_word_placeholders_in_office",
        ),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-document-authoring",),
    ),
    SurfaceLink(
        "education-booking-read-records",
        tools=(
            "microsoft_list_education_assignments",
            "microsoft_list_education_users",
            "microsoft_list_learning_course_activities",
        ),
        prompts=("microsoft-agent-operator", "microsoft-document-author"),
        skills=("microsoft-education-booking",),
    ),
    SurfaceLink(
        "endpoints-read-records",
        tools=(
            "get_intune_configuration",
            "get_windows_companion_configuration",
            "get_windows_device_health",
            "list_intune_detected_apps",
            "microsoft_list_device_compliance_policies",
            "microsoft_list_device_configurations",
        ),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "endpoints-change-records",
        tools=("remote_lock_intune_device", "shut_down_intune_device"),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-endpoint-management",),
    ),
    SurfaceLink(
        "files-read-records",
        tools=(
            "microsoft_get_drive_root_item",
            "microsoft_get_onedrive_usage_report",
            "microsoft_get_sharepoint_activity_report",
            "microsoft_list_drives",
            "microsoft_list_folder_files",
            "microsoft_list_group_drives",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "files-change-records",
        tools=("microsoft_upload_file_content",),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-files-sharepoint",),
    ),
    SurfaceLink(
        "identity-read-records-1",
        tools=(
            "microsoft_get_admin_consent_policy",
            "microsoft_get_authorization_policy",
            "microsoft_get_me",
            "microsoft_list_domain_service_configuration_records",
            "microsoft_list_entitlement_access_packages",
            "microsoft_list_group_conversations",
            "microsoft_list_group_owners",
            "microsoft_list_lifecycle_workflows",
            "microsoft_list_permission_grant_policies",
            "microsoft_list_plan_tasks",
            "microsoft_list_provisioning_logs",
            "microsoft_list_token_issuance_policies",
        ),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-read-records-2",
        tools=("microsoft_list_token_lifetime_policies", "microsoft_list_users"),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "identity-change-records",
        tools=("microsoft_create_invitation",),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-identity-directory",),
    ),
    SurfaceLink(
        "integration-read-records",
        tools=("microsoft_search_query",),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-graph-integration",),
    ),
    SurfaceLink(
        "integration-change-records",
        tools=("list_microsoft_ingestion_projection",),
        prompts=("microsoft-agent-operator", "microsoft-identity-administrator"),
        skills=("microsoft-graph-integration",),
    ),
    SurfaceLink(
        "mail-read-records",
        tools=(
            "microsoft_get_email_activity_report",
            "microsoft_get_mailbox_usage_report",
        ),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "mail-change-records",
        tools=("microsoft_create_draft_email", "microsoft_send_mail"),
        prompts=("microsoft-agent-operator", "microsoft-mailbox-coordinator"),
        skills=("microsoft-mail-operations",),
    ),
    SurfaceLink(
        "office-apps-read-records",
        tools=(
            "microsoft_get_excel_workbook",
            "microsoft_get_planner_plan",
            "microsoft_list_onenote_section_pages",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "office-apps-change-records",
        tools=("microsoft_create_excel_chart",),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-office-apps",),
    ),
    SurfaceLink(
        "power-platform-read-records",
        tools=("get_power_platform_configuration",),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-power-platform",),
    ),
    SurfaceLink(
        "print-read-records",
        tools=("microsoft_list_print_shares",),
        prompts=("microsoft-agent-operator", "microsoft-endpoint-engineer"),
        skills=("microsoft-universal-print",),
    ),
    SurfaceLink(
        "reports-read-records",
        tools=("microsoft_get_office365_active_users",),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-service-reports",),
    ),
    SurfaceLink(
        "security-read-records",
        tools=("microsoft_list_secure_scores",),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "security-change-records",
        tools=("microsoft_run_hunting_query",),
        prompts=("microsoft-agent-operator", "microsoft-security-analyst"),
        skills=("microsoft-security-operations",),
    ),
    SurfaceLink(
        "session-read-records",
        tools=("microsoft_list_accounts", "microsoft_search_tools"),
        prompts=(
            "microsoft-agent-operator",
            "microsoft-mailbox-coordinator",
            "microsoft-collaboration-manager",
            "microsoft-identity-administrator",
            "microsoft-security-analyst",
            "microsoft-endpoint-engineer",
            "microsoft-document-author",
        ),
        skills=("microsoft-session-management", "microsoft-agent-operations"),
    ),
    SurfaceLink(
        "session-change-records",
        tools=("login_power_platform", "login_windows_companion", "microsoft_logout"),
        prompts=(
            "microsoft-agent-operator",
            "microsoft-mailbox-coordinator",
            "microsoft-collaboration-manager",
            "microsoft-identity-administrator",
            "microsoft-security-analyst",
            "microsoft-endpoint-engineer",
            "microsoft-document-author",
        ),
        skills=("microsoft-session-management", "microsoft-agent-operations"),
    ),
    SurfaceLink(
        "teams-read-records",
        tools=(
            "microsoft_get_team",
            "microsoft_get_teams_user_activity",
            "microsoft_list_joined_teams",
            "microsoft_list_team_members",
        ),
        prompts=("microsoft-agent-operator", "microsoft-collaboration-manager"),
        skills=("microsoft-teams-messaging",),
    ),
)

LOCATIONS = (
    DataLocation(
        "microsoft-accounts",
        ("Account",),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_accounts",
        federation_key="username",
    ),
    DataLocation(
        "microsoft-mailboxes",
        ("Mailbox", "MailFolder", "Message", "Attachment"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_mail_messages",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-calendars",
        ("Calendar", "Event", "OnlineMeeting", "Room"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_calendar_events",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-contacts",
        ("Contact",),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_outlook_contacts",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-teams",
        ("Team", "Channel", "Chat", "ChatMessage"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_joined_teams",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-drives",
        ("Drive", "DriveItem", "Site", "SharePointList"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_drives",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-work-items",
        ("Workbook", "Notebook", "PlannerTask", "TodoTask"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_todo_task_lists",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-directory",
        ("User", "Group", "Application", "ServicePrincipal", "DirectoryRole"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_users",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-security",
        ("SecurityAlert", "SecurityIncident", "RiskyUser", "SignInLog"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_security_alerts",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-reports",
        ("ServiceHealth", "UsageReport"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_service_health",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-integration",
        ("Subscription", "ExternalConnection"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_subscriptions",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-endpoints",
        ("ManagedDevice", "CompliancePolicy", "DeviceConfiguration"),
        "MICROSOFT_GRAPH_BASE_URL",
        "list_intune_managed_devices",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-flows",
        ("Flow", "DesktopFlow", "FlowRun"),
        "MICROSOFT_GRAPH_BASE_URL",
        "list_power_automate_flows",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-office-sessions",
        ("OfficeSession", "Document"),
        "MICROSOFT_GRAPH_BASE_URL",
        "list_office_sessions",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-printers",
        ("Printer", "PrintJob"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_printers",
        federation_key="id",
    ),
    DataLocation(
        "microsoft-education-booking",
        ("EducationClass", "BookingBusiness", "LearningProvider"),
        "MICROSOFT_GRAPH_BASE_URL",
        "microsoft_list_education_classes",
        federation_key="id",
    ),
)

ROLES = (
    BusinessRole(
        "mailbox and calendar coordinator",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "mail-attachment",
            "mail-folder",
            "mail-message",
            "mail-shared-mailbox",
            "mail-read-records",
            "mail-change-records",
            "calendar-event",
            "calendar-online-meeting",
            "calendar-place",
            "calendar-specific-calendar",
            "calendar-read-records",
            "contacts-outlook-contact",
        ),
        description="Keep mail, meetings and contacts moving: triage the inbox, send what was approved, schedule meetings.",
    ),
    BusinessRole(
        "collaboration manager",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "teams-call-record",
            "teams-channel-message",
            "teams-chat",
            "teams-chat-message",
            "teams-presence",
            "teams-team-channel",
            "teams-read-records",
            "files-admin-sharepoint",
            "files-file-storage",
            "files-onedrive-file",
            "files-sharepoint-site",
            "files-site",
            "files-site-drive",
            "files-site-item",
            "files-site-list",
            "files-read-records",
            "files-change-records",
            "office-apps-excel-range",
            "office-apps-excel-table",
            "office-apps-excel-worksheet",
            "office-apps-onenote-notebook",
            "office-apps-onenote-page",
            "office-apps-planner-task",
            "office-apps-todo-task",
            "office-apps-read-records",
            "office-apps-change-records",
        ),
        description="Run team communication and shared work: Teams messages, files, workbooks, notebooks and tasks.",
    ),
    BusinessRole(
        "identity administrator",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "identity-access-review",
            "identity-agreement",
            "identity-application",
            "identity-application-password",
            "identity-conditional-access",
            "identity-delegated-admin",
            "identity-deleted-item",
            "identity-device",
            "identity-directory-object",
            "identity-directory-role",
            "identity-domain",
            "identity-group",
            "identity-group-member",
            "identity-org-branding",
            "identity-organization",
            "identity-role-assignment",
            "identity-role-definition",
            "identity-service-principal",
            "identity-read-records-1",
            "identity-read-records-2",
            "identity-change-records",
            "integration-external-connection",
            "integration-subscription",
            "integration-read-records",
            "integration-change-records",
        ),
        description="Administer directory identities and access safely, with every change approved and recorded.",
    ),
    BusinessRole(
        "security analyst",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "security-directory-audit",
            "security-risk-detection",
            "security-risky-user",
            "security-alert",
            "security-incident",
            "security-sensitivity-label",
            "security-sign-log",
            "security-subject-right",
            "security-threat-intelligence",
            "security-read-records",
            "security-change-records",
            "reports-service-health",
            "reports-service-update",
            "reports-read-records",
        ),
        description="Triage security alerts and incidents and report tenant health with evidence.",
    ),
    BusinessRole(
        "endpoint engineer",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "endpoints-intune-device",
            "endpoints-intune-managed",
            "endpoints-managed-device",
            "endpoints-window-action",
            "endpoints-read-records",
            "endpoints-change-records",
            "power-platform-power-automate",
            "power-platform-read-records",
            "print-document",
            "print-job",
            "printer",
            "print-read-records",
        ),
        description="Operate managed devices, automations and printers within the allowlists.",
    ),
    BusinessRole(
        "document author",
        tasks=(
            "session-login",
            "session-read-records",
            "session-change-records",
            "documents-powerpoint-slide",
            "documents-word-selection",
            "documents-read-records",
            "documents-change-records",
            "education-booking-booking-appointment",
            "education-booking-booking-business",
            "education-booking-education-class",
            "education-booking-education-school",
            "education-booking-learning-provider",
            "education-booking-read-records",
        ),
        description="Produce Word and PowerPoint documents and read education and booking data for them.",
    ),
)
