# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.15.0] - 2026-05-22

### Added
- Created comprehensive test suite including `test_auth_coverage.py` and `test_mcp_comprehensive.py` to cover authentication managers, token serializations, and all 36 dynamic Microsoft Graph MCP tools.
- Integrated `generate_sdd_handoff.py` and `generate_report.py` to enable automated architectural and code quality feedback.
- Documented all missing environment variables (`OIDC_CLIENT_ID`, `AUDIENCE`, etc.) in `.env.example` and the `README.md`.
- Implemented `CONCEPT:ECO-4.1` markers in source files and unit tests for 100% Concept Traceability.

### Changed
- Refactored `FALLBACK_DIR` in `microsoft_agent/auth.py` from `~/.microsoft-agent` to follow the **XDG Base Directory specification** (`~/.local/share/microsoft-agent`), including auto-migration logic from old token directories to preserve credentials.
- Updated `msgraph-sdk` dependency to `1.58.0` in package configurations to leverage the latest API stability fixes.

### Security
- Verified keyring token cache credential handling and mitigated potential credential exposure warnings.
