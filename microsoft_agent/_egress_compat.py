"""Egress (SSRF) guard for an operator-supplied/returned URL.

SDK-GAP (EH-48x, see SDK-GAPS.md #2): agent_connector_sdk has no equivalent of
agent_utilities.security.egress. This is a local, verbatim port of its
syntactic (non-DNS-resolving) half -- the only half graph_file_service.py
uses -- since it is pure stdlib with zero agent_utilities dependencies of its
own. See SDK-GAPS.md for the proposal to add this to the SDK (e.g.
agent_connector_sdk.egress) so every connector that vendors a copy of this
file can drop it.
"""

from __future__ import annotations

import ipaddress
from dataclasses import dataclass
from urllib.parse import urlparse

__all__ = ["EgressDecision", "egress_ip_is_blocked", "validate_base_url"]


@dataclass(slots=True)
class EgressDecision:
    """Outcome of an egress check."""

    allowed: bool
    reason: str = ""
    resolved_ips: tuple[str, ...] = ()


def egress_ip_is_blocked(ip: str, *, allow_loopback: bool) -> bool:
    """Whether an address is outside the public egress boundary."""
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return True  # unparseable -> block
    if addr.is_loopback:
        return not allow_loopback
    if (
        addr.is_private
        or addr.is_link_local
        or addr.is_reserved
        or addr.is_multicast
        or addr.is_unspecified
    ):
        return True
    if str(addr) in {"169.254.169.254", "fd00:ec2::254"}:  # cloud metadata service
        return True
    return False


def validate_base_url(url: str, *, allow_loopback: bool = True) -> EgressDecision:
    """Syntactic check: scheme + host present, and an IP literal host not blocked.

    Does not perform DNS resolution.
    """
    if not url or not isinstance(url, str):
        return EgressDecision(False, "empty url")
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"}:
        return EgressDecision(False, f"unsupported scheme: {parsed.scheme!r}")
    if parsed.username is not None or parsed.password is not None:
        return EgressDecision(False, "embedded URL credentials are not allowed")
    host = parsed.hostname
    if not host:
        return EgressDecision(False, "missing host")
    try:
        _ = parsed.port
    except ValueError:
        return EgressDecision(False, "invalid port")
    try:
        ipaddress.ip_address(host)
    except ValueError:
        return EgressDecision(True, "hostname (resolve to verify)")
    if egress_ip_is_blocked(host, allow_loopback=allow_loopback):
        return EgressDecision(False, "blocked IP literal")
    return EgressDecision(True, "allowed IP literal", (host,))
