"""Authenticated, durable control plane for outbound Windows companions.

The control plane exposes only the fixed endpoints used by
``WindowsCompanionClient`` and ``OutboundRelayWorker``.  Controller requests
enqueue typed actions; an authenticated device long-polls those actions over
an outbound HTTPS connection and acknowledges the typed result.  SQLite is
used for durable, bounded queue and result storage.

TLS termination is expected at a trusted reverse proxy.  A proxy-provided
mTLS certificate thumbprint is considered only when its header name is
explicitly configured; deployments must strip that header from untrusted
requests before forwarding traffic to the application.
"""

from __future__ import annotations

import asyncio
import hashlib
import hmac
import json
import ntpath
import os
import sqlite3
import ssl
from collections.abc import Mapping
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, Protocol, runtime_checkable
from urllib.error import HTTPError
from urllib.parse import quote, urlparse
from urllib.request import HTTPRedirectHandler, HTTPSHandler, Request, build_opener
from uuid import UUID, uuid4

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    HttpUrl,
    field_validator,
    model_validator,
)

from microsoft_agent.windows_companion import (
    ActionPolicy,
    CompanionActionFailure,
    CompanionActionKind,
    CompanionActionReceipt,
    CompanionActionRequest,
    CompanionActionResult,
    CompanionActionStatus,
    CompanionConnectionStatus,
    CompanionDevice,
    CompanionHealth,
    CompanionTokenProvider,
    ConfirmationRequirement,
    DeviceIdentity,
    FileListAction,
    FileReadAction,
    FileWriteAction,
    OfficeExportPdfAction,
    OfficeOpenDocumentAction,
    PowerAutomateDesktopRunAction,
    WindowsServiceStartAction,
    WindowsServiceStatusAction,
    WindowsServiceStopAction,
)
from microsoft_agent.windows_runtime import (
    OutboundRelayTransport,
    RelayActionDelivery,
    RelayPollBatch,
    WindowsRuntimeError,
)

try:  # FastAPI is supplied by the MCP/agent runtime extras.
    from fastapi import FastAPI, HTTPException, Response
    from fastapi import Request as FastAPIRequest
    from fastapi.exceptions import RequestValidationError
    from fastapi.responses import JSONResponse
except ImportError:  # pragma: no cover - minimal library-only installation
    FastAPI = None  # type: ignore[assignment,misc]
    HTTPException = None  # type: ignore[assignment,misc]
    FastAPIRequest = Any  # type: ignore[assignment,misc]
    JSONResponse = None  # type: ignore[assignment,misc]
    RequestValidationError = None  # type: ignore[assignment,misc]
    Response = None  # type: ignore[assignment,misc]

try:  # Optional production JWT signature verification dependency.
    import jwt
except ImportError:  # pragma: no cover - exercised when PyJWT is not installed
    jwt = None  # type: ignore[assignment]


class AuthenticatedPrincipal(BaseModel):
    """Identity and authorization claims from a cryptographically valid token."""

    model_config = ConfigDict(frozen=True)

    tenant_id: UUID
    subject: str = Field(min_length=1, max_length=256)
    audience: str = Field(min_length=1, max_length=512)
    issuer: str = Field(min_length=1, max_length=1024)
    client_id: str | None = Field(default=None, max_length=256)
    entra_device_id: UUID | None = None
    roles: frozenset[str] = Field(default_factory=frozenset)
    scopes: frozenset[str] = Field(default_factory=frozenset)


@runtime_checkable
class TokenValidator(Protocol):
    """Injected asynchronous bearer-token validator."""

    async def validate_token(self, token: str) -> AuthenticatedPrincipal:
        """Validate a token and return only verified claims."""


class TokenValidationError(RuntimeError):
    """A bearer token could not be securely validated."""


def _normalize_entra_v2_token_audience(value: str) -> str:
    """Return the canonical bare application GUID used by Entra v2 ``aud``."""

    candidate = value.strip()
    try:
        canonical = str(UUID(candidate))
    except (AttributeError, ValueError) as exc:
        raise ValueError(
            "Entra v2 token audience must be the API application's bare client-ID GUID"
        ) from exc
    if candidate.casefold() != canonical:
        raise ValueError(
            "Entra v2 token audience must be a canonical bare client-ID GUID"
        )
    return canonical


class StaticTokenValidator:
    """Deterministic validator for tests; never use it in production."""

    def __init__(self, tokens: Mapping[str, AuthenticatedPrincipal]) -> None:
        self._tokens = dict(tokens)

    async def validate_token(self, token: str) -> AuthenticatedPrincipal:
        principal = self._tokens.get(token)
        if principal is None:
            raise TokenValidationError("Invalid test token")
        return principal


class EntraJwtValidatorSettings(BaseModel):
    """Pinned Entra JWT validation parameters and current signing keys."""

    model_config = ConfigDict(frozen=True)

    tenant_id: UUID
    audience: str = Field(min_length=1, max_length=512)
    issuer: HttpUrl | None = None
    jwks: dict[str, Any]
    allowed_algorithms: tuple[str, ...] = ("RS256",)
    leeway_seconds: int = Field(default=60, ge=0, le=300)

    @field_validator("allowed_algorithms")
    @classmethod
    def validate_algorithms(cls, value: tuple[str, ...]) -> tuple[str, ...]:
        permitted = {"RS256", "RS384", "RS512"}
        if not value or any(item not in permitted for item in value):
            raise ValueError("Only explicit RSA SHA-2 JWT algorithms are supported")
        return value

    @field_validator("audience")
    @classmethod
    def validate_audience(cls, value: str) -> str:
        return _normalize_entra_v2_token_audience(value)

    @field_validator("jwks")
    @classmethod
    def validate_jwks(cls, value: dict[str, Any]) -> dict[str, Any]:
        keys = value.get("keys")
        if not isinstance(keys, list) or not keys:
            raise ValueError("jwks must contain at least one signing key")
        for key in keys:
            if not isinstance(key, dict) or not isinstance(key.get("kid"), str):
                raise ValueError("every JWK must have a string kid")
            if key.get("kty") != "RSA":
                raise ValueError("only RSA signing keys are supported")
            if key.get("use") not in {None, "sig"}:
                raise ValueError("JWK use must be sig when present")
        return value

    @property
    def expected_issuer(self) -> str:
        """Return the exact configured Entra v2 issuer."""

        if self.issuer is not None:
            return str(self.issuer).rstrip("/")
        return f"https://login.microsoftonline.com/{self.tenant_id}/v2.0"


class EntraJwtTokenValidator:
    """Production Entra validator with pinned JWKS, issuer, audience, and tenant."""

    def __init__(self, settings: EntraJwtValidatorSettings) -> None:
        if jwt is None:
            raise ImportError(
                "Install PyJWT with its crypto extra to validate Entra access tokens"
            )
        self.settings = settings
        self._keys: dict[str, Any] = {}
        for key in settings.jwks.get("keys", []):
            try:
                self._keys[str(key["kid"])] = jwt.algorithms.RSAAlgorithm.from_jwk(
                    json.dumps(key)
                )
            except Exception as exc:
                raise ValueError("Configured JWKS contains an invalid RSA key") from exc

    async def validate_token(self, token: str) -> AuthenticatedPrincipal:
        return await asyncio.to_thread(self._validate_sync, token)

    def _validate_sync(self, token: str) -> AuthenticatedPrincipal:
        assert jwt is not None
        try:
            header = jwt.get_unverified_header(token)
            algorithm = header.get("alg")
            key_id = header.get("kid")
            if algorithm not in self.settings.allowed_algorithms:
                raise TokenValidationError("JWT algorithm is not allowed")
            if not isinstance(key_id, str) or key_id not in self._keys:
                raise TokenValidationError("JWT signing key is not configured")
            claims = jwt.decode(
                token,
                key=self._keys[key_id],
                algorithms=list(self.settings.allowed_algorithms),
                audience=self.settings.audience,
                issuer=self.settings.expected_issuer,
                leeway=self.settings.leeway_seconds,
                options={"require": ["exp", "iss", "aud", "tid", "ver"]},
            )
        except TokenValidationError:
            raise
        except Exception as exc:
            raise TokenValidationError("Entra access token validation failed") from exc

        try:
            tenant_id = UUID(str(claims["tid"]))
        except (KeyError, TypeError, ValueError) as exc:
            raise TokenValidationError("Token tenant claim is invalid") from exc
        if tenant_id != self.settings.tenant_id:
            raise TokenValidationError("Token belongs to another tenant")
        if claims.get("ver") != "2.0":
            raise TokenValidationError("Token is not an Entra v2 access token")
        subject = claims.get("oid") or claims.get("sub")
        if not isinstance(subject, str) or not subject.strip():
            raise TokenValidationError("Token has no stable subject")
        roles_value = claims.get("roles", [])
        roles = (
            frozenset(item for item in roles_value if isinstance(item, str))
            if isinstance(roles_value, list)
            else frozenset()
        )
        scope_value = claims.get("scp", "")
        scopes = (
            frozenset(scope_value.split())
            if isinstance(scope_value, str)
            else frozenset()
        )
        if not roles and not scopes:
            raise TokenValidationError(
                "Token has neither an application role nor a delegated scope"
            )
        raw_device_id = claims.get("deviceid")
        try:
            device_id = UUID(raw_device_id) if isinstance(raw_device_id, str) else None
        except ValueError as exc:
            raise TokenValidationError(
                "Token device identity claim is invalid"
            ) from exc
        raw_client_id = claims.get("azp") or claims.get("appid")
        client_id = raw_client_id if isinstance(raw_client_id, str) else None
        return AuthenticatedPrincipal(
            tenant_id=tenant_id,
            subject=subject,
            audience=self.settings.audience,
            issuer=self.settings.expected_issuer,
            client_id=client_id,
            entra_device_id=device_id,
            roles=roles,
            scopes=scopes,
        )


class ControlPlaneDeviceRegistration(BaseModel):
    """Configured device policy and its permitted workload identities."""

    model_config = ConfigDict(frozen=True)

    device: CompanionDevice
    device_principal_subjects: frozenset[str] = Field(default_factory=frozenset)
    device_client_ids: frozenset[str] = Field(default_factory=frozenset)

    @field_validator("device_principal_subjects", "device_client_ids")
    @classmethod
    def validate_identifiers(cls, value: frozenset[str]) -> frozenset[str]:
        if any(not item.strip() or item != item.strip() for item in value):
            raise ValueError("Device identity allowlists must contain trimmed values")
        return value


class WindowsControlPlaneSettings(BaseModel):
    """Authorization and connection policy for the companion control plane."""

    model_config = ConfigDict(frozen=True)

    token_audience: str = Field(min_length=1, max_length=512)
    devices: dict[str, ControlPlaneDeviceRegistration]
    action_policies: dict[CompanionActionKind, ActionPolicy] = Field(
        default_factory=lambda: _control_plane_action_policies()
    )
    controller_roles: frozenset[str] = frozenset({"WindowsCompanion.Control"})
    controller_scopes: frozenset[str] = frozenset({"WindowsCompanion.Control"})
    controller_subjects: frozenset[str] = Field(default_factory=frozenset)
    device_roles: frozenset[str] = frozenset({"WindowsCompanion.Device"})
    device_scopes: frozenset[str] = frozenset()
    trusted_proxy_mtls_header: str | None = None
    trusted_proxy_client_hosts: frozenset[str] = Field(default_factory=frozenset)
    require_device_mtls: bool = True
    online_window_seconds: int = Field(default=90, ge=10, le=3600)
    companion_version_when_unknown: str = Field(default="not-connected", min_length=1)

    @field_validator("trusted_proxy_mtls_header")
    @classmethod
    def validate_header(cls, value: str | None) -> str | None:
        if value is None:
            return None
        if not value or any(not (char.isalnum() or char == "-") for char in value):
            raise ValueError("Trusted proxy header must be an HTTP token")
        return value

    @field_validator("token_audience")
    @classmethod
    def validate_token_audience(cls, value: str) -> str:
        return _normalize_entra_v2_token_audience(value)

    @model_validator(mode="after")
    def validate_configuration(self) -> WindowsControlPlaneSettings:
        if self.require_device_mtls and not self.trusted_proxy_mtls_header:
            raise ValueError(
                "require_device_mtls needs an explicit trusted proxy header"
            )
        if self.require_device_mtls and not self.trusted_proxy_client_hosts:
            raise ValueError(
                "require_device_mtls needs at least one trusted proxy client host"
            )
        if not (
            self.controller_roles or self.controller_scopes or self.controller_subjects
        ):
            raise ValueError("At least one controller authorization rule is required")
        if not (self.device_roles or self.device_scopes):
            raise ValueError("At least one device authorization permission is required")
        for device_id, registration in self.devices.items():
            if device_id != registration.device.identity.device_id:
                raise ValueError(
                    "Device map keys must equal configured device_id values"
                )
            missing = set(registration.device.allowed_actions) - set(
                self.action_policies
            )
            if missing:
                raise ValueError(f"Device actions have no policies: {sorted(missing)}")
        return self


class WindowsControlPlaneLimits(BaseModel):
    """Hard storage and polling bounds for the SQLite control plane."""

    model_config = ConfigDict(frozen=True)

    maximum_records: int = Field(default=50_000, ge=10, le=1_000_000)
    maximum_pending_per_device: int = Field(default=500, ge=1, le=10_000)
    maximum_submission_bytes: int = Field(default=16_777_216, ge=1024, le=16_777_216)
    maximum_result_bytes: int = Field(default=16_777_216, ge=1024, le=67_108_864)
    maximum_poll_response_bytes: int = Field(
        default=62_914_560, ge=1_048_576, le=67_108_864
    )
    maximum_database_bytes: int = Field(
        default=268_435_456, ge=1_048_576, le=10_737_418_240
    )
    completed_retention_hours: int = Field(default=168, ge=1, le=8760)
    maximum_poll_actions: int = Field(default=100, ge=1, le=100)
    maximum_poll_wait_seconds: float = Field(default=30.0, ge=0, le=60)

    @property
    def maximum_http_request_bytes(self) -> int:
        """Return the body limit including a small JSON envelope allowance."""

        return max(self.maximum_submission_bytes, self.maximum_result_bytes) + 65_536


class ActionSubmission(BaseModel):
    """Controller submission envelope accepted by the stable actions endpoint."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    request: CompanionActionRequest
    policy: ActionPolicy
    expected_device_identity: DeviceIdentity


class RelayPollRequest(BaseModel):
    """Bounded long-poll request emitted by an authenticated device."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    cursor: str | None = Field(default=None, max_length=1024)
    maximum_actions: int = Field(default=10, ge=1, le=100)
    wait_seconds: float = Field(default=30.0, ge=0, le=60)
    companion_version: str = Field(default="unknown", min_length=1, max_length=128)
    capabilities: frozenset[CompanionActionKind] = Field(default_factory=frozenset)


class _StoreConflict(RuntimeError):
    pass


class _StoreNotFound(RuntimeError):
    pass


class _StoreCapacity(RuntimeError):
    pass


class SQLiteCompanionStore:
    """Durable, bounded SQLite action queue and result store."""

    def __init__(
        self,
        database_path: str | os.PathLike[str],
        limits: WindowsControlPlaneLimits | None = None,
    ) -> None:
        self.path = Path(database_path)
        if str(self.path) == ":memory:":
            raise ValueError("The companion queue must use a durable SQLite file")
        if self.path.is_symlink():
            raise ValueError("The companion SQLite file cannot be a symlink")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.limits = limits or WindowsControlPlaneLimits()
        self._initialize()

    async def enqueue(
        self,
        device_id: str,
        submission: ActionSubmission,
    ) -> CompanionActionReceipt:
        return await asyncio.to_thread(self._enqueue_sync, device_id, submission)

    async def get_action(
        self, device_id: str, action_id: UUID
    ) -> CompanionActionResult:
        return await asyncio.to_thread(self._get_action_sync, device_id, action_id)

    async def poll(
        self,
        device_id: str,
        maximum_actions: int,
        now: datetime,
    ) -> RelayPollBatch:
        return await asyncio.to_thread(self._poll_sync, device_id, maximum_actions, now)

    async def acknowledge(
        self,
        device_id: str,
        delivery_id: str,
        result: CompanionActionResult,
    ) -> None:
        await asyncio.to_thread(self._acknowledge_sync, device_id, delivery_id, result)

    async def touch_device(
        self,
        device_id: str,
        now: datetime,
        companion_version: str,
        capabilities: frozenset[CompanionActionKind],
    ) -> None:
        await asyncio.to_thread(
            self._touch_device_sync,
            device_id,
            now,
            companion_version,
            capabilities,
        )

    async def device_state(self, device_id: str) -> dict[str, Any] | None:
        return await asyncio.to_thread(self._device_state_sync, device_id)

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=10, isolation_level=None)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=10000")
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.execute("PRAGMA journal_mode=WAL")
            page_size = int(connection.execute("PRAGMA page_size").fetchone()[0])
            maximum_pages = max(1, self.limits.maximum_database_bytes // page_size)
            connection.execute(f"PRAGMA max_page_count={maximum_pages}")
            journal_limit = min(self.limits.maximum_database_bytes // 4, 67_108_864)
            connection.execute(f"PRAGMA journal_size_limit={journal_limit}")
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS actions (
                    sequence INTEGER PRIMARY KEY AUTOINCREMENT,
                    delivery_id TEXT NOT NULL UNIQUE,
                    action_id TEXT NOT NULL UNIQUE,
                    device_id TEXT NOT NULL,
                    idempotency_key TEXT NOT NULL,
                    fingerprint TEXT NOT NULL,
                    request_json TEXT NOT NULL,
                    policy_json TEXT NOT NULL,
                    identity_json TEXT NOT NULL,
                    status TEXT NOT NULL,
                    accepted_at TEXT NOT NULL,
                    started_at TEXT,
                    completed_at TEXT,
                    expires_at TEXT NOT NULL,
                    result_json TEXT,
                    UNIQUE(device_id, idempotency_key)
                );
                CREATE INDEX IF NOT EXISTS ix_actions_device_status
                    ON actions(device_id, status, sequence);
                CREATE INDEX IF NOT EXISTS ix_actions_completed
                    ON actions(completed_at);
                CREATE TABLE IF NOT EXISTS device_state (
                    device_id TEXT PRIMARY KEY,
                    last_seen_at TEXT NOT NULL,
                    companion_version TEXT NOT NULL,
                    capabilities_json TEXT NOT NULL
                );
                """
            )
        try:
            if os.name != "nt":
                self.path.chmod(0o600)
        except OSError:
            pass

    def _enqueue_sync(
        self, device_id: str, submission: ActionSubmission
    ) -> CompanionActionReceipt:
        serialized = _canonical_json(submission.model_dump(mode="json"))
        if len(serialized.encode("utf-8")) > self.limits.maximum_submission_bytes:
            raise _StoreCapacity("Action submission exceeds the storage limit")
        fingerprint = hashlib.sha256(serialized.encode("utf-8")).hexdigest()
        now = datetime.now(UTC)
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            self._expire_and_prune(connection, now)
            existing = connection.execute(
                "SELECT * FROM actions WHERE device_id=? AND idempotency_key=?",
                (device_id, submission.request.idempotency_key),
            ).fetchone()
            if existing is not None:
                if existing["fingerprint"] != fingerprint:
                    raise _StoreConflict(
                        "Idempotency key was already used for another action"
                    )
                connection.commit()
                return self._receipt_from_row(existing)

            counts = connection.execute(
                """
                SELECT COUNT(*) AS total,
                    SUM(CASE WHEN device_id=? AND status IN ('accepted','running')
                        THEN 1 ELSE 0 END) AS pending
                FROM actions
                """,
                (device_id,),
            ).fetchone()
            if int(counts["total"] or 0) >= self.limits.maximum_records:
                raise _StoreCapacity("Companion action store reached its record limit")
            if int(counts["pending"] or 0) >= self.limits.maximum_pending_per_device:
                raise _StoreCapacity("Device action queue is full")
            page_size = int(connection.execute("PRAGMA page_size").fetchone()[0])
            page_count = int(connection.execute("PRAGMA page_count").fetchone()[0])
            if page_size * page_count >= self.limits.maximum_database_bytes:
                raise _StoreCapacity("Companion action store reached its byte limit")

            request_json = _canonical_json(submission.request.model_dump(mode="json"))
            policy_json = _canonical_json(submission.policy.model_dump(mode="json"))
            identity_json = _canonical_json(
                submission.expected_device_identity.model_dump(mode="json")
            )
            accepted_at = now.isoformat()
            try:
                connection.execute(
                    """
                    INSERT INTO actions (
                        delivery_id, action_id, device_id, idempotency_key,
                        fingerprint, request_json, policy_json, identity_json,
                        status, accepted_at, expires_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'accepted', ?, ?)
                    """,
                    (
                        str(uuid4()),
                        str(submission.request.action_id),
                        device_id,
                        submission.request.idempotency_key,
                        fingerprint,
                        request_json,
                        policy_json,
                        identity_json,
                        accepted_at,
                        submission.request.expires_at.isoformat(),
                    ),
                )
            except sqlite3.IntegrityError as exc:
                raise _StoreConflict(
                    "Action ID or idempotency key already exists"
                ) from exc
            row = connection.execute(
                "SELECT * FROM actions WHERE action_id=?",
                (str(submission.request.action_id),),
            ).fetchone()
            connection.commit()
            assert row is not None
            return self._receipt_from_row(row)

    def _get_action_sync(
        self, device_id: str, action_id: UUID
    ) -> CompanionActionResult:
        now = datetime.now(UTC)
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            self._expire_and_prune(connection, now)
            row = connection.execute(
                "SELECT * FROM actions WHERE device_id=? AND action_id=?",
                (device_id, str(action_id)),
            ).fetchone()
            connection.commit()
        if row is None:
            raise _StoreNotFound("Action was not found")
        if row["result_json"]:
            return CompanionActionResult.model_validate_json(row["result_json"])
        return CompanionActionResult(
            action_id=action_id,
            device_id=device_id,
            status=CompanionActionStatus(row["status"]),
            started_at=_parse_optional_datetime(row["started_at"]),
        )

    def _poll_sync(
        self, device_id: str, maximum_actions: int, now: datetime
    ) -> RelayPollBatch:
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            self._expire_and_prune(connection, now)
            rows = connection.execute(
                """
                SELECT * FROM actions
                WHERE device_id=? AND status IN ('accepted','running')
                ORDER BY sequence ASC LIMIT ?
                """,
                (device_id, maximum_actions),
            ).fetchall()
            bounded_rows: list[sqlite3.Row] = []
            response_bytes = 1024
            for row in rows:
                estimated = (
                    len(row["request_json"].encode("utf-8"))
                    + len(row["policy_json"].encode("utf-8"))
                    + len(row["identity_json"].encode("utf-8"))
                    + 2048
                )
                if (
                    bounded_rows
                    and response_bytes + estimated
                    > self.limits.maximum_poll_response_bytes
                ):
                    break
                bounded_rows.append(row)
                response_bytes += estimated
            rows = bounded_rows
            for row in rows:
                if row["status"] == CompanionActionStatus.ACCEPTED.value:
                    connection.execute(
                        "UPDATE actions SET status='running', started_at=? WHERE sequence=?",
                        (now.isoformat(), row["sequence"]),
                    )
            connection.commit()
        deliveries = tuple(
            RelayActionDelivery(
                delivery_id=row["delivery_id"],
                device_id=device_id,
                expected_device_identity=DeviceIdentity.model_validate_json(
                    row["identity_json"]
                ),
                request=CompanionActionRequest.model_validate_json(row["request_json"]),
                policy=ActionPolicy.model_validate_json(row["policy_json"]),
            )
            for row in rows
        )
        cursor = str(rows[-1]["sequence"]) if rows else None
        return RelayPollBatch(cursor=cursor, deliveries=deliveries)

    def _acknowledge_sync(
        self,
        device_id: str,
        delivery_id: str,
        result: CompanionActionResult,
    ) -> None:
        terminal = {
            CompanionActionStatus.SUCCEEDED,
            CompanionActionStatus.FAILED,
            CompanionActionStatus.REJECTED,
            CompanionActionStatus.EXPIRED,
            CompanionActionStatus.CANCELED,
        }
        if result.status not in terminal:
            raise _StoreConflict("Only terminal action results can be acknowledged")
        result_json = _canonical_json(result.model_dump(mode="json"))
        if len(result_json.encode("utf-8")) > self.limits.maximum_result_bytes:
            raise _StoreCapacity("Action result exceeds the storage limit")
        now = datetime.now(UTC)
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM actions WHERE device_id=? AND delivery_id=?",
                (device_id, delivery_id),
            ).fetchone()
            if row is None:
                raise _StoreNotFound("Relay delivery was not found")
            if (
                row["action_id"] != str(result.action_id)
                or result.device_id != device_id
            ):
                raise _StoreConflict("Acknowledgement identity did not match delivery")
            if row["result_json"]:
                if not hmac.compare_digest(row["result_json"], result_json):
                    raise _StoreConflict("Delivery already has a different result")
                connection.commit()
                return
            connection.execute(
                """
                UPDATE actions SET status=?, completed_at=?, result_json=?
                WHERE sequence=?
                """,
                (result.status.value, now.isoformat(), result_json, row["sequence"]),
            )
            connection.commit()

    def _touch_device_sync(
        self,
        device_id: str,
        now: datetime,
        companion_version: str,
        capabilities: frozenset[CompanionActionKind],
    ) -> None:
        capabilities_json = _canonical_json(sorted(item.value for item in capabilities))
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO device_state (
                    device_id, last_seen_at, companion_version, capabilities_json
                ) VALUES (?, ?, ?, ?)
                ON CONFLICT(device_id) DO UPDATE SET
                    last_seen_at=excluded.last_seen_at,
                    companion_version=excluded.companion_version,
                    capabilities_json=excluded.capabilities_json
                """,
                (device_id, now.isoformat(), companion_version, capabilities_json),
            )

    def _device_state_sync(self, device_id: str) -> dict[str, Any] | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM device_state WHERE device_id=?", (device_id,)
            ).fetchone()
        if row is None:
            return None
        return {
            "last_seen_at": datetime.fromisoformat(row["last_seen_at"]),
            "companion_version": row["companion_version"],
            "capabilities": frozenset(
                CompanionActionKind(item)
                for item in json.loads(row["capabilities_json"])
            ),
        }

    def _expire_and_prune(self, connection: sqlite3.Connection, now: datetime) -> None:
        rows = connection.execute(
            """
            SELECT * FROM actions
            WHERE status IN ('accepted','running') AND expires_at<=?
            """,
            (now.isoformat(),),
        ).fetchall()
        for row in rows:
            result = CompanionActionResult(
                action_id=UUID(row["action_id"]),
                device_id=row["device_id"],
                status=CompanionActionStatus.EXPIRED,
                started_at=_parse_optional_datetime(row["started_at"]),
                completed_at=now,
                error=CompanionActionFailure(
                    code="expired",
                    message="Action expired before device acknowledgement",
                    retryable=False,
                ),
            )
            connection.execute(
                """
                UPDATE actions SET status='expired', completed_at=?, result_json=?
                WHERE sequence=?
                """,
                (
                    now.isoformat(),
                    _canonical_json(result.model_dump(mode="json")),
                    row["sequence"],
                ),
            )
        cutoff = now - timedelta(hours=self.limits.completed_retention_hours)
        connection.execute(
            """
            DELETE FROM actions
            WHERE completed_at IS NOT NULL AND completed_at<?
            """,
            (cutoff.isoformat(),),
        )

    @staticmethod
    def _receipt_from_row(row: sqlite3.Row) -> CompanionActionReceipt:
        return CompanionActionReceipt(
            action_id=UUID(row["action_id"]),
            device_id=row["device_id"],
            status=CompanionActionStatus(row["status"]),
            accepted_at=datetime.fromisoformat(row["accepted_at"]),
            status_url=(
                f"/v1/devices/{quote(row['device_id'], safe='')}/actions/"
                f"{row['action_id']}"
            ),
        )


class _PayloadTooLarge(RuntimeError):
    pass


class _RequestSizeLimitMiddleware:
    """Bound request bodies before FastAPI allocates or validates their JSON."""

    def __init__(self, app: Any, maximum_bytes: int) -> None:
        self._app = app
        self._maximum_bytes = maximum_bytes

    async def __call__(self, scope: Any, receive: Any, send: Any) -> None:
        if scope.get("type") != "http":
            await self._app(scope, receive, send)
            return
        headers = {key.lower(): value for key, value in scope.get("headers", [])}
        raw_length = headers.get(b"content-length")
        if raw_length is not None:
            try:
                if int(raw_length) > self._maximum_bytes:
                    await self._reject(send)
                    return
            except ValueError:
                await self._reject(send)
                return
        received = 0
        response_started = False

        async def limited_receive() -> Any:
            nonlocal received
            message = await receive()
            if message.get("type") == "http.request":
                received += len(message.get("body", b""))
                if received > self._maximum_bytes:
                    raise _PayloadTooLarge
            return message

        async def tracked_send(message: Any) -> None:
            nonlocal response_started
            if message.get("type") == "http.response.start":
                response_started = True
            await send(message)

        try:
            await self._app(scope, limited_receive, tracked_send)
        except _PayloadTooLarge:
            if not response_started:
                await self._reject(send)

    @staticmethod
    async def _reject(send: Any) -> None:
        body = b'{"detail":"Request body is too large"}'
        await send(
            {
                "type": "http.response.start",
                "status": 413,
                "headers": [
                    (b"content-type", b"application/json"),
                    (b"content-length", str(len(body)).encode("ascii")),
                ],
            }
        )
        await send({"type": "http.response.body", "body": body})


def create_windows_control_plane_app(
    settings: WindowsControlPlaneSettings,
    store: SQLiteCompanionStore,
    token_validator: TokenValidator,
) -> Any:
    """Create the fixed FastAPI control-plane application."""

    if (
        FastAPI is None
        or HTTPException is None
        or Response is None
        or JSONResponse is None
        or RequestValidationError is None
    ):
        raise ImportError("Install FastAPI to host the Windows control plane")
    app = FastAPI(
        title="Microsoft Agent Windows Companion Control Plane",
        docs_url=None,
        redoc_url=None,
        openapi_url=None,
    )
    configured_devices = dict(settings.devices)
    configured_policies = dict(settings.action_policies)
    app.add_middleware(
        _RequestSizeLimitMiddleware,
        maximum_bytes=store.limits.maximum_http_request_bytes,
    )

    @app.exception_handler(RequestValidationError)
    async def invalid_request(_request: Any, _error: Any) -> Any:
        return JSONResponse(
            status_code=422,
            content={"detail": "Request body or path parameters are invalid"},
        )

    async def principal(request: FastAPIRequest) -> AuthenticatedPrincipal:
        authorization = request.headers.get("authorization", "")
        scheme, separator, token = authorization.partition(" ")
        if (
            not separator
            or scheme.casefold() != "bearer"
            or not token
            or any(char.isspace() for char in token)
        ):
            raise HTTPException(
                status_code=401,
                detail="A bearer access token is required",
                headers={"WWW-Authenticate": "Bearer"},
            )
        try:
            verified = await token_validator.validate_token(token)
        except Exception as exc:
            raise HTTPException(
                status_code=401,
                detail="Bearer access token validation failed",
                headers={"WWW-Authenticate": "Bearer"},
            ) from exc
        if not hmac.compare_digest(verified.audience, settings.token_audience):
            raise HTTPException(status_code=401, detail="Token audience is invalid")
        return verified

    def registration(device_id: str) -> ControlPlaneDeviceRegistration:
        configured = configured_devices.get(device_id)
        if configured is None:
            raise HTTPException(status_code=404, detail="Device was not found")
        if not configured.device.enabled:
            raise HTTPException(status_code=403, detail="Device is disabled")
        return configured

    async def authorize_controller(
        request: FastAPIRequest, device_id: str
    ) -> ControlPlaneDeviceRegistration:
        configured = registration(device_id)
        verified = await principal(request)
        if verified.tenant_id != configured.device.identity.tenant_id:
            raise HTTPException(status_code=403, detail="Tenant is not authorized")
        authorized = (
            verified.subject in settings.controller_subjects
            or bool(verified.roles & settings.controller_roles)
            or bool(verified.scopes & settings.controller_scopes)
        )
        if not authorized:
            raise HTTPException(status_code=403, detail="Controller is not authorized")
        header_device = request.headers.get("x-microsoft-agent-device-id")
        if header_device is None or not hmac.compare_digest(header_device, device_id):
            raise HTTPException(
                status_code=403, detail="Device header did not match path"
            )
        return configured

    async def authorize_device(
        request: FastAPIRequest, device_id: str
    ) -> ControlPlaneDeviceRegistration:
        configured = registration(device_id)
        verified = await principal(request)
        identity = configured.device.identity
        if verified.tenant_id != identity.tenant_id:
            raise HTTPException(status_code=403, detail="Tenant is not authorized")
        if not (
            verified.roles & settings.device_roles
            or verified.scopes & settings.device_scopes
        ):
            raise HTTPException(
                status_code=403, detail="Device token permission is not authorized"
            )
        if (
            verified.entra_device_id is not None
            and verified.entra_device_id != identity.entra_device_id
        ):
            raise HTTPException(
                status_code=403, detail="Token device claim did not match"
            )
        identity_matches = (
            verified.entra_device_id == identity.entra_device_id
            or verified.subject in configured.device_principal_subjects
            or (
                verified.client_id is not None
                and verified.client_id in configured.device_client_ids
            )
        )
        if not identity_matches:
            raise HTTPException(
                status_code=403, detail="Device identity is not authorized"
            )
        header_device = request.headers.get("x-microsoft-agent-device-id")
        if header_device is None or not hmac.compare_digest(header_device, device_id):
            raise HTTPException(
                status_code=403, detail="Device header did not match path"
            )
        if settings.require_device_mtls:
            assert settings.trusted_proxy_mtls_header is not None
            client_host = request.client.host if request.client is not None else None
            if client_host not in settings.trusted_proxy_client_hosts:
                raise HTTPException(
                    status_code=403,
                    detail="mTLS identity header did not come from a trusted proxy",
                )
            actual_thumbprint = _normalize_thumbprint(
                request.headers.get(settings.trusted_proxy_mtls_header, "")
            )
            if not hmac.compare_digest(
                actual_thumbprint, identity.certificate_thumbprint
            ):
                raise HTTPException(
                    status_code=403, detail="Client certificate did not match device"
                )
        return configured

    @app.get("/v1/devices/{device_id}/health", response_model=CompanionHealth)
    async def get_health(device_id: str, request: FastAPIRequest) -> CompanionHealth:
        configured = await authorize_controller(request, device_id)
        state = await store.device_state(device_id)
        now = datetime.now(UTC)
        online = bool(
            state
            and state["last_seen_at"]
            >= now - timedelta(seconds=settings.online_window_seconds)
        )
        return CompanionHealth(
            identity=configured.device.identity,
            status=(
                CompanionConnectionStatus.ONLINE
                if online
                else CompanionConnectionStatus.OFFLINE
            ),
            authenticated=state is not None,
            outbound_connected=online,
            last_seen_at=state["last_seen_at"] if state else now,
            companion_version=(
                state["companion_version"]
                if state
                else settings.companion_version_when_unknown
            ),
            capabilities=(
                state["capabilities"]
                if state
                else frozenset(configured.device.allowed_actions)
            ),
        )

    @app.post(
        "/v1/devices/{device_id}/actions",
        response_model=CompanionActionReceipt,
        status_code=202,
    )
    async def submit_action(
        device_id: str, submission: ActionSubmission, request: FastAPIRequest
    ) -> CompanionActionReceipt:
        configured = await authorize_controller(request, device_id)
        try:
            _validate_submission(configured_policies, configured.device, submission)
            return await store.enqueue(device_id, submission)
        except _StoreConflict as exc:
            raise HTTPException(
                status_code=409, detail="action state conflict"
            ) from exc
        except _StoreCapacity as exc:
            raise HTTPException(
                status_code=429, detail="action capacity exceeded"
            ) from exc
        except ValueError as exc:
            raise HTTPException(
                status_code=403, detail="action denied by policy"
            ) from exc

    @app.get(
        "/v1/devices/{device_id}/actions/{action_id}",
        response_model=CompanionActionResult,
    )
    async def get_action(
        device_id: str, action_id: UUID, request: FastAPIRequest
    ) -> CompanionActionResult:
        await authorize_controller(request, device_id)
        try:
            return await store.get_action(device_id, action_id)
        except _StoreNotFound as exc:
            raise HTTPException(status_code=404, detail="action not found") from exc

    @app.post("/v1/devices/{device_id}/relay/poll", response_model=RelayPollBatch)
    async def poll_actions(
        device_id: str, body: RelayPollRequest, request: FastAPIRequest
    ) -> RelayPollBatch:
        configured = await authorize_device(request, device_id)
        allowed_maximum = min(
            body.maximum_actions, settings_for_store(store).maximum_poll_actions
        )
        allowed_wait = min(
            body.wait_seconds, settings_for_store(store).maximum_poll_wait_seconds
        )
        capabilities = frozenset(body.capabilities) & frozenset(
            configured.device.allowed_actions
        )
        deadline = asyncio.get_running_loop().time() + allowed_wait
        while True:
            now = datetime.now(UTC)
            await store.touch_device(
                device_id, now, body.companion_version, capabilities
            )
            batch = await store.poll(device_id, allowed_maximum, now)
            if batch.deliveries or asyncio.get_running_loop().time() >= deadline:
                if batch.cursor is None:
                    batch = batch.model_copy(update={"cursor": body.cursor})
                return batch
            await asyncio.sleep(
                min(0.25, max(0, deadline - asyncio.get_running_loop().time()))
            )

    @app.post(
        "/v1/devices/{device_id}/relay/actions/{delivery_id}/ack",
        status_code=204,
        response_model=None,
    )
    async def acknowledge_action(
        device_id: str,
        delivery_id: str,
        result: CompanionActionResult,
        request: FastAPIRequest,
    ) -> Any:
        await authorize_device(request, device_id)
        try:
            await store.acknowledge(device_id, delivery_id, result)
        except _StoreNotFound as exc:
            raise HTTPException(status_code=404, detail="delivery not found") from exc
        except _StoreConflict as exc:
            raise HTTPException(
                status_code=409, detail="delivery state conflict"
            ) from exc
        except _StoreCapacity as exc:
            raise HTTPException(
                status_code=413, detail="result capacity exceeded"
            ) from exc
        return Response(status_code=204)

    return app


def settings_for_store(store: SQLiteCompanionStore) -> WindowsControlPlaneLimits:
    """Return the store's validated hard limits."""

    return store.limits


class _NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args: Any, **kwargs: Any) -> None:  # noqa: ARG002
        return None


class HttpOutboundRelayTransport(OutboundRelayTransport):
    """HTTPS, bearer-authenticated transport for the outbound device worker."""

    def __init__(
        self,
        control_plane_url: str,
        token_audience: str,
        token_provider: CompanionTokenProvider,
        *,
        timeout_seconds: float = 35,
        companion_version: str = "microsoft-agent",
        capabilities: frozenset[CompanionActionKind] = frozenset(),
        client_certificate_path: str | os.PathLike[str] | None = None,
        client_private_key_path: str | os.PathLike[str] | None = None,
        private_key_password: str | None = None,
        ca_bundle_path: str | os.PathLike[str] | None = None,
        maximum_response_bytes: int = 67_108_864,
    ) -> None:
        parsed = urlparse(control_plane_url)
        if (
            parsed.scheme != "https"
            or not parsed.netloc
            or parsed.username
            or parsed.password
            or parsed.path not in {"", "/"}
            or parsed.query
            or parsed.fragment
        ):
            raise ValueError("Control-plane URL must be an HTTPS origin")
        if bool(client_certificate_path) != bool(client_private_key_path):
            raise ValueError("Both client certificate and private key are required")
        self._base_url = control_plane_url.rstrip("/")
        self._audience = token_audience
        self._token_provider = token_provider
        self._timeout = timeout_seconds
        self._version = companion_version
        self._capabilities = capabilities
        if not 1024 <= maximum_response_bytes <= 268_435_456:
            raise ValueError("maximum_response_bytes is outside the safe range")
        self._maximum_response_bytes = maximum_response_bytes
        context = ssl.create_default_context(
            cafile=str(ca_bundle_path) if ca_bundle_path else None
        )
        if client_certificate_path and client_private_key_path:
            context.load_cert_chain(
                certfile=str(client_certificate_path),
                keyfile=str(client_private_key_path),
                password=private_key_password,
            )
        self._opener = build_opener(_NoRedirect(), HTTPSHandler(context=context))

    async def poll(
        self,
        identity: DeviceIdentity,
        *,
        cursor: str | None,
        maximum_actions: int,
        wait_seconds: float,
    ) -> RelayPollBatch:
        payload = RelayPollRequest(
            cursor=cursor,
            maximum_actions=maximum_actions,
            wait_seconds=wait_seconds,
            companion_version=self._version,
            capabilities=self._capabilities,
        )
        body = await self._post(
            identity,
            f"/v1/devices/{quote(identity.device_id, safe='')}/relay/poll",
            payload.model_dump_json().encode("utf-8"),
            expected={200},
        )
        try:
            return RelayPollBatch.model_validate_json(body)
        except (ValueError, json.JSONDecodeError) as exc:
            raise WindowsRuntimeError(
                "invalid_relay_response",
                "Control plane returned an invalid poll response",
            ) from exc

    async def acknowledge(
        self,
        identity: DeviceIdentity,
        delivery_id: str,
        result: CompanionActionResult,
    ) -> None:
        safe_delivery = quote(delivery_id, safe="")
        await self._post(
            identity,
            (
                f"/v1/devices/{quote(identity.device_id, safe='')}/relay/actions/"
                f"{safe_delivery}/ack"
            ),
            result.model_dump_json().encode("utf-8"),
            expected={204},
        )

    async def _post(
        self,
        identity: DeviceIdentity,
        path: str,
        body: bytes,
        *,
        expected: set[int],
    ) -> bytes:
        try:
            token = await asyncio.wait_for(
                self._token_provider.get_token(self._audience), timeout=self._timeout
            )
        except Exception as exc:
            raise WindowsRuntimeError(
                "relay_authentication_failed",
                f"Relay token acquisition failed: {type(exc).__name__}",
            ) from exc
        if not token or not token.strip():
            raise WindowsRuntimeError(
                "relay_authentication_failed", "Relay token provider returned no token"
            )

        def send() -> tuple[int, bytes]:
            request = Request(
                f"{self._base_url}{path}",
                data=body,
                method="POST",
                headers={
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {token}",
                    "X-Microsoft-Agent-Device-ID": identity.device_id,
                },
            )
            try:
                with self._opener.open(request, timeout=self._timeout) as response:
                    return (
                        int(response.status),
                        response.read(self._maximum_response_bytes + 1),
                    )
            except HTTPError as exc:
                return int(exc.code), exc.read(
                    min(self._maximum_response_bytes + 1, 65_536)
                )

        try:
            status, response_body = await asyncio.wait_for(
                asyncio.to_thread(send), timeout=self._timeout
            )
        except Exception as exc:
            raise WindowsRuntimeError(
                "relay_transport_failed",
                f"Relay HTTPS request failed: {type(exc).__name__}",
            ) from exc
        if status not in expected:
            raise WindowsRuntimeError(
                "relay_http_error", f"Control plane returned HTTP {status}"
            )
        if len(response_body) > self._maximum_response_bytes:
            raise WindowsRuntimeError(
                "relay_response_too_large",
                "Control-plane response exceeded the configured byte limit",
            )
        return response_body


def _validate_submission(
    action_policies: Mapping[CompanionActionKind, ActionPolicy],
    device: CompanionDevice,
    submission: ActionSubmission,
) -> None:
    if submission.expected_device_identity != device.identity:
        raise ValueError("Expected device identity did not match configuration")
    action = submission.request.action
    kind = CompanionActionKind(action.kind)
    if kind not in device.allowed_actions:
        raise ValueError("Action is not allowed for this device")
    local_policy = action_policies.get(kind)
    if local_policy is None or submission.policy != local_policy:
        raise ValueError("Submitted policy did not match control-plane policy")
    now = datetime.now(UTC)
    if submission.request.expires_at <= now:
        raise ValueError("Action request has expired")
    if isinstance(action, (FileListAction, FileReadAction, FileWriteAction)):
        _require_logical_path(device, action.path)
    elif isinstance(action, OfficeOpenDocumentAction):
        _require_logical_path(device, action.document_path)
    elif isinstance(action, OfficeExportPdfAction):
        _require_logical_path(device, action.source_path)
        _require_logical_path(device, action.output_path)
    elif isinstance(action, PowerAutomateDesktopRunAction):
        if action.flow_name.casefold() not in {
            item.casefold() for item in device.allowed_desktop_flows
        }:
            raise ValueError("Desktop flow is not allowlisted")
    elif isinstance(
        action,
        (
            WindowsServiceStatusAction,
            WindowsServiceStartAction,
            WindowsServiceStopAction,
        ),
    ) and action.service_name.casefold() not in {
        item.casefold() for item in device.allowed_services
    }:
        raise ValueError("Windows service is not allowlisted")
    if local_policy.requires_confirmation:
        evidence = submission.request.confirmation
        if evidence is None:
            raise ValueError("Action requires confirmation")
        if evidence.action_kind is not kind:
            raise ValueError("Confirmation belongs to another action kind")
        if evidence.confirmed_at > now or evidence.expires_at <= now:
            raise ValueError("Confirmation is not currently valid")


def _require_logical_path(device: CompanionDevice, value: str) -> None:
    candidate = ntpath.normcase(ntpath.normpath(value.strip()))
    if not ntpath.isabs(candidate):
        raise ValueError("File path must be absolute")
    if candidate.startswith(("\\\\?\\", "\\\\.\\", "\\??\\")):
        raise ValueError("Windows device namespace paths are forbidden")
    for root in device.allowed_file_roots:
        normalized_root = ntpath.normcase(ntpath.normpath(root))
        try:
            if ntpath.commonpath((candidate, normalized_root)) == normalized_root:
                relative = ntpath.relpath(candidate, normalized_root)
                if relative == "." or all(
                    part not in {"", ".", ".."} and ":" not in part
                    for part in relative.split("\\")
                ):
                    return
        except ValueError:
            continue
    raise ValueError("File path is outside the device allowlist")


def _control_plane_action_policies() -> dict[CompanionActionKind, ActionPolicy]:
    read = ActionPolicy(
        confirmation=ConfirmationRequirement.NONE,
        rationale="Read-only device metadata",
    )
    sensitive = ActionPolicy(
        confirmation=ConfirmationRequirement.ALWAYS,
        rationale="May expose user or file content",
    )
    change = ActionPolicy(
        confirmation=ConfirmationRequirement.WHEN_DESTRUCTIVE,
        destructive=True,
        rationale="Changes state on the Windows device",
    )
    return {
        CompanionActionKind.SYSTEM_INVENTORY: read,
        CompanionActionKind.FILE_LIST: sensitive,
        CompanionActionKind.FILE_READ: sensitive,
        CompanionActionKind.FILE_WRITE: change,
        CompanionActionKind.OFFICE_OPEN_DOCUMENT: change,
        CompanionActionKind.OFFICE_EXPORT_PDF: change,
        CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN: change,
        CompanionActionKind.WINDOWS_SERVICE_STATUS: read,
        CompanionActionKind.WINDOWS_SERVICE_START: change,
        CompanionActionKind.WINDOWS_SERVICE_STOP: change,
        CompanionActionKind.NOTIFICATION_SHOW: change,
        CompanionActionKind.CLIPBOARD_READ_TEXT: sensitive,
        CompanionActionKind.CLIPBOARD_WRITE_TEXT: change,
    }


def _canonical_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def _parse_optional_datetime(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None


def _normalize_thumbprint(value: str) -> str:
    cleaned = value.replace(":", "").replace(" ", "").upper()
    return cleaned if len(cleaned) in {40, 64} and cleaned.isalnum() else ""


__all__ = [
    "ActionSubmission",
    "AuthenticatedPrincipal",
    "ControlPlaneDeviceRegistration",
    "EntraJwtTokenValidator",
    "EntraJwtValidatorSettings",
    "HttpOutboundRelayTransport",
    "RelayPollRequest",
    "SQLiteCompanionStore",
    "StaticTokenValidator",
    "TokenValidationError",
    "TokenValidator",
    "WindowsControlPlaneLimits",
    "WindowsControlPlaneSettings",
    "create_windows_control_plane_app",
]
