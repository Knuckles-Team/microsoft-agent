"""Validate the checked-in Office add-in manifest without a network service."""

from __future__ import annotations

import argparse
import re
import sys
import uuid
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urlsplit

OFFICE_NAMESPACE = "http://schemas.microsoft.com/office/appforoffice/1.1"
XSI_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance"
VERSION = re.compile(r"^[0-9]+(?:\.[0-9]+){3}$")


class ManifestError(ValueError):
    """A stable manifest validation failure."""


def _tag(name: str) -> str:
    """Return a fully qualified Office manifest element name."""

    return f"{{{OFFICE_NAMESPACE}}}{name}"


def _one(root: ET.Element, name: str) -> ET.Element:
    values = root.findall(_tag(name))
    if len(values) != 1:
        raise ManifestError(f"manifest must contain exactly one {name}")
    return values[0]


def _text(root: ET.Element, name: str) -> str:
    value = (_one(root, name).text or "").strip()
    if not value:
        raise ManifestError(f"manifest {name} must not be empty")
    return value


def _default_value(root: ET.Element, name: str) -> str:
    value = (_one(root, name).get("DefaultValue") or "").strip()
    if not value:
        raise ManifestError(f"manifest {name} DefaultValue must not be empty")
    return value


def _https_url(value: str, field: str) -> None:
    parsed = urlsplit(value)
    if (
        parsed.scheme != "https"
        or not parsed.hostname
        or parsed.username is not None
        or parsed.password is not None
        or parsed.query
        or parsed.fragment
    ):
        raise ManifestError(f"manifest {field} must be an authority-only HTTPS URL")


def validate_manifest(path: Path) -> None:
    """Validate well-formedness and the provider's supported manifest contract."""

    payload = path.read_bytes()
    if len(payload) > 256 * 1024:
        raise ManifestError("manifest exceeds its size boundary")
    upper = payload.upper()
    if b"<!DOCTYPE" in upper or b"<!ENTITY" in upper:
        raise ManifestError("manifest must not declare a DTD or entity")
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ManifestError("manifest XML is not well formed") from exc

    if root.tag != _tag("OfficeApp"):
        raise ManifestError("manifest root namespace or element is unsupported")
    if root.get(f"{{{XSI_NAMESPACE}}}type") != "TaskPaneApp":
        raise ManifestError("manifest must declare the TaskPaneApp type")

    try:
        uuid.UUID(_text(root, "Id"))
    except ValueError as exc:
        raise ManifestError("manifest Id must be a UUID") from exc
    if not VERSION.fullmatch(_text(root, "Version")):
        raise ManifestError("manifest Version must contain four numeric components")

    for field in ("ProviderName", "DefaultLocale", "Permissions"):
        _text(root, field)
    for field in ("DisplayName", "Description"):
        _default_value(root, field)
    for field in ("IconUrl", "HighResolutionIconUrl", "SupportUrl"):
        _https_url(_default_value(root, field), field)

    domains = _one(root, "AppDomains").findall(_tag("AppDomain"))
    if not domains:
        raise ManifestError("manifest must allow at least one application domain")
    for domain in domains:
        _https_url((domain.text or "").strip(), "AppDomain")

    hosts = {
        (host.get("Name") or "").strip()
        for host in _one(root, "Hosts").findall(_tag("Host"))
    }
    if hosts != {"Document", "Presentation"}:
        raise ManifestError("manifest hosts must be Document and Presentation")

    source = _one(_one(root, "DefaultSettings"), "SourceLocation")
    _https_url((source.get("DefaultValue") or "").strip(), "SourceLocation")
    if _text(root, "Permissions") != "ReadWriteDocument":
        raise ManifestError("manifest permission must be ReadWriteDocument")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    arguments = parser.parse_args(argv)
    try:
        validate_manifest(arguments.manifest)
    except (ManifestError, OSError) as exc:
        print(f"manifest validation failed: {exc}", file=sys.stderr)
        return 1
    print("manifest validation: passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
