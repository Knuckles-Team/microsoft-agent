"""Offline Office manifest validation contract."""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VALIDATOR = ROOT / "office_addin" / "scripts" / "validate_manifest.py"
MANIFEST = ROOT / "office_addin" / "manifest.xml"


def _validate(path: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(VALIDATOR), str(path)],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
        timeout=10,
    )


def test_checked_in_office_manifest_passes_offline_contract() -> None:
    result = _validate(MANIFEST)

    assert result.returncode == 0
    assert result.stdout.strip() == "manifest validation: passed"


def test_office_manifest_rejects_doctype(tmp_path: Path) -> None:
    invalid = tmp_path / "manifest.xml"
    invalid.write_text(
        '<!DOCTYPE OfficeApp [<!ENTITY payload "unsafe">]><OfficeApp>&payload;</OfficeApp>',
        encoding="utf-8",
    )

    result = _validate(invalid)

    assert result.returncode == 1
    assert result.stderr.strip() == (
        "manifest validation failed: manifest must not declare a DTD or entity"
    )


def test_office_manifest_rejects_insecure_origin(tmp_path: Path) -> None:
    invalid = tmp_path / "manifest.xml"
    invalid.write_text(
        MANIFEST.read_text(encoding="utf-8").replace(
            "https://localhost:3000/assets/icon.png",
            "http://localhost:3000/assets/icon.png",
            1,
        ),
        encoding="utf-8",
    )

    result = _validate(invalid)

    assert result.returncode == 1
    assert result.stderr.strip() == (
        "manifest validation failed: manifest IconUrl must be an authority-only HTTPS URL"
    )
