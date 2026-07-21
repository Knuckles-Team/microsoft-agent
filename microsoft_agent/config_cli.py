"""Offline, secret-free Microsoft Agent configuration validation CLI."""

from __future__ import annotations

import argparse
import json
from typing import Any

from microsoft_agent.document_service import get_document_capabilities
from microsoft_agent.integration_config import get_integration_settings
from microsoft_agent.settings import get_settings


def configuration_report() -> dict[str, Any]:
    """Build a sanitized readiness report without acquiring any token."""

    settings = get_settings()
    integrations = get_integration_settings()
    identity_error: str | None = None
    if settings.is_configured:
        try:
            settings.require_identity()
        except ValueError:
            identity_error = "invalid_identity_configuration"

    power = integrations.power_platform
    companion = integrations.windows_companion
    intune = integrations.intune
    return {
        "identity": {
            "configured": settings.is_configured and identity_error is None,
            "mode": settings.authentication_mode.value,
            "tenant_configured": bool(settings.tenant_id),
            "client_configured": bool(settings.client_id),
            "validation_error": identity_error,
            "permission_profiles": [
                profile.value for profile in settings.permission_profiles
            ],
        },
        "policy": {
            "writes_enabled": settings.allow_write_tools,
            "destructive_enabled": settings.allow_destructive_tools,
            "enabled_tool_groups": list(settings.enabled_tool_groups),
        },
        "documents": get_document_capabilities().model_dump(mode="json"),
        "power_platform": {
            "configured": power is not None,
            "named_flow_count": len(power.named_flows) if power else 0,
        },
        "windows_companion": {
            "configured": companion is not None,
            "configured_device_count": len(companion.devices) if companion else 0,
        },
        "intune": {
            "configured": intune is not None,
            "allowed_device_count": len(intune.allowed_device_ids) if intune else 0,
        },
        "office_addin": {
            "allowed_origin_count": len(settings.office_addin_origins),
        },
        "ingestion": {
            "pseudonymization_key_configured": bool(
                settings.ingestion_pseudonymization_key
            )
        },
    }


def main(argv: list[str] | None = None) -> int:
    """Validate configuration and print only non-secret readiness metadata."""

    parser = argparse.ArgumentParser(
        description="Validate Microsoft Agent configuration without authenticating."
    )
    parser.add_argument(
        "--require-identity",
        action="store_true",
        help="Exit non-zero unless identity coordinates/files are configured.",
    )
    args = parser.parse_args(argv)
    try:
        report = configuration_report()
    except ValueError:
        print(json.dumps({"valid": False, "error": "invalid_configuration"}, indent=2))
        return 2
    report["valid"] = True
    print(json.dumps(report, indent=2, sort_keys=True))
    if args.require_identity and not report["identity"]["configured"]:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
