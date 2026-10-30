"""The control-plane client TLS context comes from the SDK TLS profile."""

from __future__ import annotations

import ast
import ssl
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pytest
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.x509.oid import NameOID
from scion_sdk.tls.errors import TransportSecurityError

from microsoft_agent import windows_runtime
from microsoft_agent.windows_runtime import build_client_tls_context

ROOT = Path(__file__).resolve().parents[1]


def _identity(directory: Path) -> tuple[Path, Path]:
    key = ec.generate_private_key(ec.SECP256R1())
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "companion")])
    now = datetime.now(UTC)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - timedelta(minutes=1))
        .not_valid_after(now + timedelta(days=1))
        .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
        .sign(key, hashes.SHA256())
    )
    cert_path, key_path = directory / "cert.pem", directory / "key.pem"
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    key_path.write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption(),
        )
    )
    return cert_path, key_path


def test_client_context_verifies_peer_and_uses_instance_ca(tmp_path: Path) -> None:
    """CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001: profile-built verified context."""
    cert, key = _identity(tmp_path)
    context = build_client_tls_context(
        ca_bundle_path=cert,
        client_certificate_path=cert,
        client_private_key_path=key,
    )
    assert context.verify_mode == ssl.CERT_REQUIRED
    assert context.check_hostname is True
    assert context.minimum_version >= ssl.TLSVersion.TLSv1_2
    subjects = [entry.get("subject") for entry in context.get_ca_certs()]
    assert ((("commonName", "companion"),),) in subjects


def test_client_context_uses_the_sdk_profile(monkeypatch: pytest.MonkeyPatch) -> None:
    """CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001: resolution goes through the SDK."""
    seen: dict[str, object] = {}
    real = windows_runtime.resolve_tls_profile

    def spy(service: str, **kwargs: object):
        seen.update(service=service, **kwargs)
        return real(service, **kwargs)

    monkeypatch.setattr(windows_runtime, "resolve_tls_profile", spy)
    build_client_tls_context()
    assert seen["service"] == "microsoft"


def test_missing_ca_bundle_fails_closed(tmp_path: Path) -> None:
    """CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R001: no silent fallback to default trust."""
    with pytest.raises(TransportSecurityError):
        build_client_tls_context(ca_bundle_path=tmp_path / "absent.pem")


def test_no_hand_built_tls_context_in_package() -> None:
    """CONN-MICROSOFT-AGENT-TRANSPORT-TLS-R002: only the SDK builds SSL contexts."""
    offenders: list[str] = []
    for path in (ROOT / "microsoft_agent").rglob("*.py"):
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
            if (
                isinstance(node, ast.Call)
                and isinstance(node.func, ast.Attribute)
                and node.func.attr
                in {
                    "create_default_context",
                    "SSLContext",
                    "_create_unverified_context",
                }
            ):
                offenders.append(f"{path.relative_to(ROOT)}:{node.lineno}")
            if isinstance(node, ast.keyword) and node.arg == "verify":
                if isinstance(node.value, ast.Constant) and node.value.value is False:
                    offenders.append(f"{path.relative_to(ROOT)}:{node.value.lineno}")
    assert not offenders
