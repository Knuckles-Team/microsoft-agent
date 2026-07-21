"""Typed, short-lived bridge between MCP tools and a paired Office.js task pane.

The bridge deliberately supports only the Word and PowerPoint operations modeled
in this module.  An MCP caller enqueues a validated command, while an Office.js
task pane long-polls over HTTPS and posts a typed result.  Pairing and session
secrets are random, expire, are stored only as SHA-256 digests, and never grant
access to Microsoft Graph or another Office session.

State is intentionally process-local.  Restarting the MCP server revokes every
pairing and session, which is safer than silently persisting bearer credentials.
Deploy the HTTP transport with a single worker or replace ``OfficeBridgeStore``
with a shared implementation before horizontally scaling the service.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import secrets
from collections import defaultdict, deque
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import UTC, datetime, timedelta
from enum import StrEnum
from typing import Annotated, Any, Literal
from uuid import UUID, uuid4

from fastmcp import FastMCP
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    SecretStr,
    TypeAdapter,
    ValidationError,
    field_validator,
    model_validator,
)
from starlette.requests import Request
from starlette.responses import JSONResponse, Response

from microsoft_agent.settings import get_settings

MAX_HTTP_BODY_BYTES = 262_144
PAIRING_TTL = timedelta(minutes=5)
SESSION_TTL = timedelta(hours=8)
SESSION_IDLE_TTL = timedelta(minutes=15)
COMMAND_TTL = timedelta(minutes=2)
RESULT_TTL = timedelta(minutes=10)


class OfficeHost(StrEnum):
    """Office hosts supported by the live bridge."""

    WORD = "Word"
    POWERPOINT = "PowerPoint"


class WordInsertMode(StrEnum):
    """Supported Word selection insertion modes."""

    REPLACE = "Replace"
    START = "Start"
    END = "End"


class OfficeCommandKind(StrEnum):
    """Complete allowlist of executable Office.js commands."""

    WORD_READ_SELECTION = "word.read_selection"
    WORD_WRITE_SELECTION = "word.write_selection"
    WORD_REPLACE_PLACEHOLDERS = "word.replace_placeholders"
    POWERPOINT_LIST_SLIDES = "powerpoint.list_slides"
    POWERPOINT_ADD_SLIDE = "powerpoint.add_slide"
    POWERPOINT_DELETE_SLIDE = "powerpoint.delete_slide"
    POWERPOINT_INSERT_TEXT_BOX = "powerpoint.insert_text_box"


class OfficeCommandState(StrEnum):
    """Lifecycle state visible to MCP callers."""

    QUEUED = "queued"
    DELIVERED = "delivered"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    EXPIRED = "expired"


class OfficeRequirementSupport(BaseModel):
    """One Office.js requirement-set capability reported by the task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    requirement_set: Literal["WordApi", "PowerPointApi"]
    version: str = Field(min_length=1, max_length=32, pattern=r"^[0-9]+(?:\.[0-9]+)*$")
    supported: bool


class OfficeCapabilityReport(BaseModel):
    """Bounded host information supplied while redeeming a pairing secret."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    host: OfficeHost
    platform: str = Field(min_length=1, max_length=128)
    office_version: str = Field(min_length=1, max_length=64)
    requirements: tuple[OfficeRequirementSupport, ...] = Field(max_length=32)


class WordReadSelectionCommand(BaseModel):
    """Read the current selection from an open Word document."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_READ_SELECTION] = (
        OfficeCommandKind.WORD_READ_SELECTION
    )


class WordWriteSelectionCommand(BaseModel):
    """Write bounded text at the current Word selection."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_WRITE_SELECTION] = (
        OfficeCommandKind.WORD_WRITE_SELECTION
    )
    content: str = Field(max_length=100_000)
    mode: WordInsertMode = WordInsertMode.REPLACE


class WordReplacePlaceholdersCommand(BaseModel):
    """Replace a literal, single-line placeholder in the open Word document."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_REPLACE_PLACEHOLDERS] = (
        OfficeCommandKind.WORD_REPLACE_PLACEHOLDERS
    )
    placeholder: str = Field(min_length=1, max_length=500)
    replacement: str = Field(max_length=100_000)

    @field_validator("placeholder")
    @classmethod
    def validate_placeholder(cls, value: str) -> str:
        if "\r" in value or "\n" in value:
            raise ValueError("Placeholder must be a single line")
        return value


class PowerPointListSlidesCommand(BaseModel):
    """List slides in the open PowerPoint presentation."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_LIST_SLIDES] = (
        OfficeCommandKind.POWERPOINT_LIST_SLIDES
    )


class PowerPointAddSlideCommand(BaseModel):
    """Append a slide to the open PowerPoint presentation."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_ADD_SLIDE] = (
        OfficeCommandKind.POWERPOINT_ADD_SLIDE
    )


class PowerPointDeleteSlideCommand(BaseModel):
    """Delete a one-based slide number from the open presentation."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_DELETE_SLIDE] = (
        OfficeCommandKind.POWERPOINT_DELETE_SLIDE
    )
    slide_number: int = Field(ge=1, le=10_000)


class PowerPointInsertTextBoxCommand(BaseModel):
    """Insert a bounded text box into one PowerPoint slide."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_INSERT_TEXT_BOX] = (
        OfficeCommandKind.POWERPOINT_INSERT_TEXT_BOX
    )
    slide_number: int = Field(ge=1, le=10_000)
    text: str = Field(min_length=1, max_length=100_000)
    left: float = Field(ge=0, le=10_000)
    top: float = Field(ge=0, le=10_000)
    width: float = Field(gt=0, le=10_000)
    height: float = Field(gt=0, le=10_000)


OfficeCommandPayload = Annotated[
    WordReadSelectionCommand
    | WordWriteSelectionCommand
    | WordReplacePlaceholdersCommand
    | PowerPointListSlidesCommand
    | PowerPointAddSlideCommand
    | PowerPointDeleteSlideCommand
    | PowerPointInsertTextBoxCommand,
    Field(discriminator="kind"),
]


class OfficeCommandEnvelope(BaseModel):
    """Command delivered to exactly one paired Office task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    command_id: UUID
    session_id: UUID
    created_at: datetime
    expires_at: datetime
    payload: OfficeCommandPayload


class SlideSummary(BaseModel):
    """Safe slide metadata returned from PowerPoint."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    number: int = Field(ge=1, le=10_000)
    id: str = Field(min_length=1, max_length=512)


class WordSelectionResult(BaseModel):
    """Result of reading the current Word selection."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_READ_SELECTION]
    text: str = Field(max_length=100_000)


class WordWriteResult(BaseModel):
    """Acknowledgement for a Word selection write."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_WRITE_SELECTION]
    applied: Literal[True]


class WordPlaceholderResult(BaseModel):
    """Count returned by a Word placeholder replacement."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.WORD_REPLACE_PLACEHOLDERS]
    replacements: int = Field(ge=0, le=100_000)


class PowerPointSlidesResult(BaseModel):
    """Bounded slide list returned by PowerPoint."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_LIST_SLIDES]
    slides: tuple[SlideSummary, ...] = Field(max_length=10_000)


class PowerPointAddSlideResult(BaseModel):
    """Metadata for a newly appended slide."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_ADD_SLIDE]
    slide: SlideSummary


class PowerPointDeleteSlideResult(BaseModel):
    """Acknowledgement for a PowerPoint slide deletion."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_DELETE_SLIDE]
    deleted: Literal[True]


class PowerPointInsertTextBoxResult(BaseModel):
    """Shape identifier for an inserted PowerPoint text box."""

    model_config = ConfigDict(extra="forbid", frozen=True)
    kind: Literal[OfficeCommandKind.POWERPOINT_INSERT_TEXT_BOX]
    shape_id: str = Field(min_length=1, max_length=512)


OfficeCommandResult = Annotated[
    WordSelectionResult
    | WordWriteResult
    | WordPlaceholderResult
    | PowerPointSlidesResult
    | PowerPointAddSlideResult
    | PowerPointDeleteSlideResult
    | PowerPointInsertTextBoxResult,
    Field(discriminator="kind"),
]


class OfficeCommandError(BaseModel):
    """Sanitized failure returned by an Office task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    code: Literal["capability", "validation", "office", "unexpected"]
    message: str = Field(min_length=1, max_length=1_000)


class OfficeCommandSucceeded(BaseModel):
    """Typed successful result posted by the task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    command_id: UUID
    status: Literal["succeeded"]
    kind: OfficeCommandKind
    result: OfficeCommandResult

    @model_validator(mode="after")
    def validate_result_kind(self) -> OfficeCommandSucceeded:
        if self.result.kind != self.kind:
            raise ValueError("Result kind does not match outcome kind")
        return self


class OfficeCommandFailed(BaseModel):
    """Typed failure posted by the task pane without a stack trace."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    command_id: UUID
    status: Literal["failed"]
    kind: OfficeCommandKind
    error: OfficeCommandError


OfficeCommandOutcome = Annotated[
    OfficeCommandSucceeded | OfficeCommandFailed,
    Field(discriminator="status"),
]
OUTCOME_ADAPTER = TypeAdapter(OfficeCommandOutcome)


class OfficeCommandReceipt(BaseModel):
    """Current lifecycle and optional outcome for an MCP-enqueued command."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    command_id: UUID
    session_id: UUID
    kind: OfficeCommandKind
    state: OfficeCommandState
    created_at: datetime
    expires_at: datetime
    outcome: OfficeCommandOutcome | None = None


class OfficePairingGrant(BaseModel):
    """One-time secret an operator pastes into a visible Office task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    pairing_id: UUID
    pairing_token: str = Field(min_length=32, max_length=256, repr=False)
    expected_host: OfficeHost
    label: str
    expires_at: datetime


class OfficePairSessionRequest(BaseModel):
    """Request made by an Office task pane to redeem a pairing secret."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    pairing_token: SecretStr
    capabilities: OfficeCapabilityReport


class OfficeSessionGrant(BaseModel):
    """Short-lived bearer credential returned only to the paired task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    session_id: UUID
    session_token: str = Field(min_length=32, max_length=256, repr=False)
    host: OfficeHost
    label: str
    expires_at: datetime


class OfficeSessionSummary(BaseModel):
    """Non-secret paired-session metadata visible to MCP callers."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    session_id: UUID
    host: OfficeHost
    label: str
    connected_at: datetime
    last_seen_at: datetime
    expires_at: datetime
    capabilities: OfficeCapabilityReport
    queued_commands: int = Field(ge=0)


class OfficePollRequest(BaseModel):
    """Bounded long-poll request from a paired task pane."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    wait_seconds: float = Field(default=5, ge=0, le=5)


class OfficePollResponse(BaseModel):
    """At most one command returned to avoid parallel document mutations."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    command: OfficeCommandEnvelope | None


class OfficeBridgeError(RuntimeError):
    """Sanitized domain error carrying a stable machine-readable code."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


@dataclass(slots=True)
class _PairingRecord:
    pairing_id: UUID
    token_digest: bytes
    expected_host: OfficeHost
    label: str
    expires_at: datetime


@dataclass(slots=True)
class _SessionRecord:
    session_id: UUID
    token_digest: bytes
    host: OfficeHost
    label: str
    connected_at: datetime
    last_seen_at: datetime
    expires_at: datetime
    capabilities: OfficeCapabilityReport
    queue: deque[UUID] = field(default_factory=deque)
    available: asyncio.Event = field(default_factory=asyncio.Event)


@dataclass(slots=True)
class _CommandRecord:
    envelope: OfficeCommandEnvelope
    state: OfficeCommandState
    outcome: OfficeCommandOutcome | None = None
    result_expires_at: datetime | None = None
    completed: asyncio.Event = field(default_factory=asyncio.Event)


def _utcnow() -> datetime:
    return datetime.now(UTC)


def _token_digest(token: str) -> bytes:
    return hashlib.sha256(token.encode("utf-8")).digest()


def _host_for_kind(kind: OfficeCommandKind) -> OfficeHost:
    if kind.value.startswith("word."):
        return OfficeHost.WORD
    return OfficeHost.POWERPOINT


class OfficeBridgeStore:
    """Concurrency-safe, bounded, in-memory pairing and command store."""

    def __init__(
        self,
        *,
        clock: Callable[[], datetime] = _utcnow,
        max_pairings: int = 64,
        max_sessions: int = 32,
        max_commands_per_session: int = 64,
    ) -> None:
        if min(max_pairings, max_sessions, max_commands_per_session) < 1:
            raise ValueError("Office bridge capacity limits must be positive")
        self._clock = clock
        self._max_pairings = max_pairings
        self._max_sessions = max_sessions
        self._max_commands_per_session = max_commands_per_session
        self._lock = asyncio.Lock()
        self._pairings: dict[bytes, _PairingRecord] = {}
        self._sessions: dict[UUID, _SessionRecord] = {}
        self._session_tokens: dict[bytes, UUID] = {}
        self._commands: dict[UUID, _CommandRecord] = {}
        self._session_commands: dict[UUID, set[UUID]] = defaultdict(set)

    def _now(self) -> datetime:
        now = self._clock()
        if now.tzinfo is None:
            raise ValueError("Office bridge clock must return an aware datetime")
        return now.astimezone(UTC)

    def _cleanup_locked(self, now: datetime) -> None:
        for digest, pairing in tuple(self._pairings.items()):
            if pairing.expires_at <= now:
                self._pairings.pop(digest, None)

        for session_id, session in tuple(self._sessions.items()):
            if (
                session.expires_at <= now
                or session.last_seen_at + SESSION_IDLE_TTL <= now
            ):
                self._expire_session_locked(session_id, now)

        for command_id, command in tuple(self._commands.items()):
            if (
                command.state
                in {OfficeCommandState.QUEUED, OfficeCommandState.DELIVERED}
                and command.envelope.expires_at <= now
            ):
                command.state = OfficeCommandState.EXPIRED
                command.result_expires_at = now + RESULT_TTL
                command.completed.set()
            if (
                command.result_expires_at is not None
                and command.result_expires_at <= now
            ):
                self._remove_command_locked(command_id)

    def _expire_session_locked(self, session_id: UUID, now: datetime) -> None:
        session = self._sessions.pop(session_id, None)
        if session is None:
            return
        self._session_tokens.pop(session.token_digest, None)
        session.available.set()
        for command_id in self._session_commands.get(session_id, set()):
            command = self._commands.get(command_id)
            if command is not None and command.state in {
                OfficeCommandState.QUEUED,
                OfficeCommandState.DELIVERED,
            }:
                command.state = OfficeCommandState.EXPIRED
                command.result_expires_at = now + RESULT_TTL
                command.completed.set()

    def _remove_command_locked(self, command_id: UUID) -> None:
        command = self._commands.pop(command_id, None)
        if command is None:
            return
        session_commands = self._session_commands.get(command.envelope.session_id)
        if session_commands is not None:
            session_commands.discard(command_id)
            if not session_commands:
                self._session_commands.pop(command.envelope.session_id, None)

    def _session_for_token_locked(self, token: str, now: datetime) -> _SessionRecord:
        if not token or len(token) > 256:
            raise OfficeBridgeError("invalid_session", "Office session is invalid")
        session_id = self._session_tokens.get(_token_digest(token))
        session = self._sessions.get(session_id) if session_id is not None else None
        if session is None:
            raise OfficeBridgeError("invalid_session", "Office session is invalid")
        session.last_seen_at = now
        return session

    async def create_pairing(
        self, expected_host: OfficeHost, label: str
    ) -> OfficePairingGrant:
        """Create a one-time, five-minute pairing secret."""

        normalized_label = label.strip()
        if not normalized_label or len(normalized_label) > 128:
            raise OfficeBridgeError(
                "invalid_label", "Office session label must be 1 to 128 characters"
            )
        token = secrets.token_urlsafe(32)
        digest = _token_digest(token)
        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            if len(self._pairings) >= self._max_pairings:
                raise OfficeBridgeError(
                    "capacity", "Office pairing capacity has been reached"
                )
            pairing = _PairingRecord(
                pairing_id=uuid4(),
                token_digest=digest,
                expected_host=expected_host,
                label=normalized_label,
                expires_at=now + PAIRING_TTL,
            )
            self._pairings[digest] = pairing
        return OfficePairingGrant(
            pairing_id=pairing.pairing_id,
            pairing_token=token,
            expected_host=pairing.expected_host,
            label=pairing.label,
            expires_at=pairing.expires_at,
        )

    async def pair_session(
        self, request: OfficePairSessionRequest
    ) -> OfficeSessionGrant:
        """Consume one pairing secret and issue a hashed, expiring session token."""

        token = request.pairing_token.get_secret_value()
        digest = _token_digest(token)
        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            pairing = self._pairings.get(digest)
            if pairing is None:
                raise OfficeBridgeError(
                    "invalid_pairing", "Pairing secret is invalid or expired"
                )
            if pairing.expected_host != request.capabilities.host:
                raise OfficeBridgeError(
                    "host_mismatch", "Pairing secret was issued for another Office host"
                )
            if len(self._sessions) >= self._max_sessions:
                raise OfficeBridgeError(
                    "capacity", "Office session capacity has been reached"
                )
            self._pairings.pop(digest, None)
            session_token = secrets.token_urlsafe(32)
            session_digest = _token_digest(session_token)
            session = _SessionRecord(
                session_id=uuid4(),
                token_digest=session_digest,
                host=pairing.expected_host,
                label=pairing.label,
                connected_at=now,
                last_seen_at=now,
                expires_at=now + SESSION_TTL,
                capabilities=request.capabilities,
            )
            self._sessions[session.session_id] = session
            self._session_tokens[session_digest] = session.session_id
        return OfficeSessionGrant(
            session_id=session.session_id,
            session_token=session_token,
            host=session.host,
            label=session.label,
            expires_at=session.expires_at,
        )

    async def list_sessions(self) -> tuple[OfficeSessionSummary, ...]:
        """List live sessions without returning bearer credentials."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            summaries = [
                self._summary_locked(session) for session in self._sessions.values()
            ]
        return tuple(sorted(summaries, key=lambda item: item.connected_at))

    def _summary_locked(self, session: _SessionRecord) -> OfficeSessionSummary:
        queued = sum(
            1
            for command_id in self._session_commands.get(session.session_id, set())
            if (command := self._commands.get(command_id)) is not None
            and command.state is OfficeCommandState.QUEUED
        )
        return OfficeSessionSummary(
            session_id=session.session_id,
            host=session.host,
            label=session.label,
            connected_at=session.connected_at,
            last_seen_at=session.last_seen_at,
            expires_at=session.expires_at,
            capabilities=session.capabilities,
            queued_commands=queued,
        )

    async def disconnect(self, session_token: str) -> None:
        """Revoke a task-pane session immediately."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            session = self._session_for_token_locked(session_token, now)
            self._expire_session_locked(session.session_id, now)

    async def enqueue(
        self, session_id: UUID, payload: OfficeCommandPayload
    ) -> OfficeCommandReceipt:
        """Queue one allowlisted command for a live host-compatible session."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            session = self._sessions.get(session_id)
            if session is None:
                raise OfficeBridgeError(
                    "session_not_found", "Office session is not connected"
                )
            if session.host != _host_for_kind(payload.kind):
                raise OfficeBridgeError(
                    "host_mismatch", "Command does not match the paired Office host"
                )
            if (
                len(self._session_commands[session_id])
                >= self._max_commands_per_session
            ):
                raise OfficeBridgeError(
                    "capacity",
                    "Office command capacity has been reached for this session",
                )
            envelope = OfficeCommandEnvelope(
                command_id=uuid4(),
                session_id=session_id,
                created_at=now,
                expires_at=now + COMMAND_TTL,
                payload=payload,
            )
            command = _CommandRecord(
                envelope=envelope,
                state=OfficeCommandState.QUEUED,
            )
            self._commands[envelope.command_id] = command
            self._session_commands[session_id].add(envelope.command_id)
            session.queue.append(envelope.command_id)
            session.available.set()
            return self._receipt_locked(command)

    async def poll(
        self, session_token: str, wait_seconds: float = 5
    ) -> OfficeCommandEnvelope | None:
        """Long-poll and deliver at most one command without unsafe redelivery."""

        deadline = asyncio.get_running_loop().time() + max(0, min(wait_seconds, 5))
        while True:
            now = self._now()
            async with self._lock:
                self._cleanup_locked(now)
                session = self._session_for_token_locked(session_token, now)
                while session.queue:
                    command_id = session.queue.popleft()
                    command = self._commands.get(command_id)
                    if (
                        command is None
                        or command.state is not OfficeCommandState.QUEUED
                    ):
                        continue
                    command.state = OfficeCommandState.DELIVERED
                    if not session.queue:
                        session.available.clear()
                    return command.envelope
                session.available.clear()
                available = session.available
            remaining = deadline - asyncio.get_running_loop().time()
            if remaining <= 0:
                return None
            try:
                await asyncio.wait_for(available.wait(), timeout=remaining)
            except TimeoutError:
                return None

    async def complete(
        self, session_token: str, outcome: OfficeCommandOutcome
    ) -> OfficeCommandReceipt:
        """Accept one typed terminal result from the command's paired session."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            session = self._session_for_token_locked(session_token, now)
            command = self._commands.get(outcome.command_id)
            if command is None or command.envelope.session_id != session.session_id:
                raise OfficeBridgeError(
                    "command_not_found", "Office command was not found"
                )
            if command.state in {
                OfficeCommandState.SUCCEEDED,
                OfficeCommandState.FAILED,
            }:
                if command.outcome == outcome:
                    return self._receipt_locked(command)
                raise OfficeBridgeError(
                    "conflict", "Office command already has a result"
                )
            if command.state is not OfficeCommandState.DELIVERED:
                raise OfficeBridgeError(
                    "conflict", "Office command is not awaiting a result"
                )
            if command.envelope.payload.kind != outcome.kind:
                raise OfficeBridgeError(
                    "kind_mismatch", "Office result does not match the command kind"
                )
            command.outcome = outcome
            command.state = (
                OfficeCommandState.SUCCEEDED
                if outcome.status == "succeeded"
                else OfficeCommandState.FAILED
            )
            command.result_expires_at = now + RESULT_TTL
            command.completed.set()
            return self._receipt_locked(command)

    async def get_command(self, command_id: UUID) -> OfficeCommandReceipt:
        """Return current command state until its bounded result retention expires."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            command = self._commands.get(command_id)
            if command is None:
                raise OfficeBridgeError(
                    "command_not_found", "Office command was not found"
                )
            return self._receipt_locked(command)

    async def wait_for_result(
        self, command_id: UUID, wait_seconds: float
    ) -> OfficeCommandReceipt:
        """Wait a bounded time for a command, returning pending state on timeout."""

        now = self._now()
        async with self._lock:
            self._cleanup_locked(now)
            command = self._commands.get(command_id)
            if command is None:
                raise OfficeBridgeError(
                    "command_not_found", "Office command was not found"
                )
            if command.state in {
                OfficeCommandState.SUCCEEDED,
                OfficeCommandState.FAILED,
                OfficeCommandState.EXPIRED,
            }:
                return self._receipt_locked(command)
            completed = command.completed
        try:
            await asyncio.wait_for(
                completed.wait(), timeout=max(0, min(wait_seconds, 55))
            )
        except TimeoutError:
            pass
        return await self.get_command(command_id)

    @staticmethod
    def _receipt_locked(command: _CommandRecord) -> OfficeCommandReceipt:
        return OfficeCommandReceipt(
            command_id=command.envelope.command_id,
            session_id=command.envelope.session_id,
            kind=command.envelope.payload.kind,
            state=command.state,
            created_at=command.envelope.created_at,
            expires_at=command.envelope.expires_at,
            outcome=command.outcome,
        )


_store: OfficeBridgeStore | None = None


def get_office_bridge_store() -> OfficeBridgeStore:
    """Return the process-local Office bridge store."""

    global _store
    if _store is None:
        _store = OfficeBridgeStore()
    return _store


def clear_office_bridge_store() -> None:
    """Revoke all in-memory pairings and sessions; primarily useful in tests."""

    global _store
    _store = None


def _cors_headers(request: Request) -> dict[str, str]:
    origin = request.headers.get("origin")
    if origin is None or origin not in get_settings().office_addin_origins:
        raise OfficeBridgeError("origin_denied", "Office add-in origin is not allowed")
    return {
        "Access-Control-Allow-Origin": origin,
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Authorization, Content-Type",
        "Access-Control-Max-Age": "600",
        "Cache-Control": "no-store",
        "Vary": "Origin",
    }


def _bearer_token(request: Request) -> str:
    authorization = request.headers.get("authorization", "")
    scheme, separator, token = authorization.partition(" ")
    if separator != " " or scheme.casefold() != "bearer" or not token:
        raise OfficeBridgeError("invalid_session", "Office session is invalid")
    if (
        len(token) > 256
        or token.strip() != token
        or any(char.isspace() for char in token)
    ):
        raise OfficeBridgeError("invalid_session", "Office session is invalid")
    return token


async def _read_json(request: Request) -> Any:
    declared_length = request.headers.get("content-length")
    if declared_length is not None:
        try:
            if int(declared_length) > MAX_HTTP_BODY_BYTES:
                raise OfficeBridgeError("body_too_large", "Request body is too large")
        except ValueError as exc:
            raise OfficeBridgeError(
                "invalid_request", "Content-Length is invalid"
            ) from exc
    body = await request.body()
    if len(body) > MAX_HTTP_BODY_BYTES:
        raise OfficeBridgeError("body_too_large", "Request body is too large")
    try:
        return json.loads(body)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise OfficeBridgeError(
            "invalid_request", "Request body must be valid JSON"
        ) from exc


def _http_error(error: OfficeBridgeError, headers: dict[str, str]) -> JSONResponse:
    statuses = {
        "invalid_pairing": 401,
        "invalid_session": 401,
        "origin_denied": 403,
        "session_not_found": 404,
        "command_not_found": 404,
        "host_mismatch": 409,
        "kind_mismatch": 409,
        "conflict": 409,
        "capacity": 429,
        "body_too_large": 413,
    }
    return JSONResponse(
        {"error": error.code, "message": "Office bridge request failed"},
        status_code=statuses.get(error.code, 400),
        headers=headers,
    )


async def _dispatch_and_wait(
    store: OfficeBridgeStore,
    session_id: UUID,
    payload: OfficeCommandPayload,
    wait_timeout_seconds: float,
) -> dict[str, Any]:
    receipt = await store.enqueue(session_id, payload)
    if wait_timeout_seconds > 0:
        receipt = await store.wait_for_result(receipt.command_id, wait_timeout_seconds)
    return receipt.model_dump(mode="json")


def register_office_bridge(
    mcp: FastMCP, store: OfficeBridgeStore | None = None
) -> None:
    """Register typed MCP tools and exact-origin Office.js relay routes."""

    bridge = store or get_office_bridge_store()

    @mcp.tool(
        name="create_office_pairing",
        description=(
            "Create a five-minute, one-time pairing secret for a visible Word or "
            "PowerPoint task pane. Give the secret only to the intended user."
        ),
        tags={"documents", "office", "pairing"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def create_pairing(expected_host: OfficeHost, label: str) -> dict[str, Any]:
        """Create an expiring pairing grant bound to one Office host and label."""

        grant = await bridge.create_pairing(expected_host, label)
        return grant.model_dump(mode="json")

    @mcp.tool(
        name="list_office_sessions",
        description="List paired, currently connected Office task panes without secrets.",
        tags={"documents", "office"},
        annotations={"readOnlyHint": True},
    )
    async def list_sessions() -> list[dict[str, Any]]:
        """Return live Office session IDs, hosts, labels, and capabilities."""

        sessions = await bridge.list_sessions()
        return [item.model_dump(mode="json") for item in sessions]

    @mcp.tool(
        name="get_office_command_result",
        description="Get the state or retained typed result of an Office bridge command.",
        tags={"documents", "office"},
        annotations={"readOnlyHint": True},
    )
    async def get_command_result(command_id: UUID) -> dict[str, Any]:
        """Read one queued, delivered, completed, failed, or expired command."""

        return (await bridge.get_command(command_id)).model_dump(mode="json")

    @mcp.tool(
        name="get_word_selection_from_office",
        description="Read the current selection from a paired, open Word task pane.",
        tags={"documents", "word", "office"},
        annotations={"readOnlyHint": True},
    )
    async def get_word_selection(
        session_id: UUID,
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a typed Word selection read and optionally await its result."""

        return await _dispatch_and_wait(
            bridge, session_id, WordReadSelectionCommand(), wait_timeout_seconds
        )

    @mcp.tool(
        name="write_word_selection_in_office",
        description="Write bounded text at the selection in a paired, open Word task pane.",
        tags={"documents", "word", "office"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def write_word_selection(
        session_id: UUID,
        content: str = Field(max_length=100_000),
        mode: WordInsertMode = WordInsertMode.REPLACE,
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a typed Word selection write and optionally await its result."""

        return await _dispatch_and_wait(
            bridge,
            session_id,
            WordWriteSelectionCommand(content=content, mode=mode),
            wait_timeout_seconds,
        )

    @mcp.tool(
        name="replace_word_placeholders_in_office",
        description="Replace a literal placeholder throughout a paired, open Word document.",
        tags={"documents", "word", "office"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def replace_word_placeholders(
        session_id: UUID,
        placeholder: str = Field(min_length=1, max_length=500),
        replacement: str = Field(max_length=100_000),
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a bounded Word placeholder replacement."""

        return await _dispatch_and_wait(
            bridge,
            session_id,
            WordReplacePlaceholdersCommand(
                placeholder=placeholder, replacement=replacement
            ),
            wait_timeout_seconds,
        )

    @mcp.tool(
        name="list_powerpoint_slides_in_office",
        description="List slides in a paired, open PowerPoint task pane.",
        tags={"documents", "powerpoint", "office"},
        annotations={"readOnlyHint": True},
    )
    async def list_powerpoint_slides(
        session_id: UUID,
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a typed PowerPoint slide-list request."""

        return await _dispatch_and_wait(
            bridge, session_id, PowerPointListSlidesCommand(), wait_timeout_seconds
        )

    @mcp.tool(
        name="add_powerpoint_slide_in_office",
        description="Append a slide in a paired, open PowerPoint task pane.",
        tags={"documents", "powerpoint", "office"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def add_powerpoint_slide(
        session_id: UUID,
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a typed PowerPoint append-slide request."""

        return await _dispatch_and_wait(
            bridge, session_id, PowerPointAddSlideCommand(), wait_timeout_seconds
        )

    @mcp.tool(
        name="delete_powerpoint_slide_in_office",
        description="Delete one numbered slide in a paired, open PowerPoint task pane.",
        tags={"documents", "powerpoint", "office"},
        annotations={"readOnlyHint": False, "destructiveHint": True},
    )
    async def delete_powerpoint_slide(
        session_id: UUID,
        slide_number: int = Field(ge=1, le=10_000),
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a policy-classified destructive PowerPoint slide deletion."""

        return await _dispatch_and_wait(
            bridge,
            session_id,
            PowerPointDeleteSlideCommand(slide_number=slide_number),
            wait_timeout_seconds,
        )

    @mcp.tool(
        name="insert_powerpoint_text_box_in_office",
        description="Insert a positioned text box in a paired, open PowerPoint task pane.",
        tags={"documents", "powerpoint", "office"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def insert_powerpoint_text_box(
        session_id: UUID,
        slide_number: int = Field(ge=1, le=10_000),
        text: str = Field(min_length=1, max_length=100_000),
        left: float = Field(ge=0, le=10_000),
        top: float = Field(ge=0, le=10_000),
        width: float = Field(gt=0, le=10_000),
        height: float = Field(gt=0, le=10_000),
        wait_timeout_seconds: float = Field(default=20, ge=0, le=55),
    ) -> dict[str, Any]:
        """Enqueue a bounded PowerPoint text-box insertion."""

        return await _dispatch_and_wait(
            bridge,
            session_id,
            PowerPointInsertTextBoxCommand(
                slide_number=slide_number,
                text=text,
                left=left,
                top=top,
                width=width,
                height=height,
            ),
            wait_timeout_seconds,
        )

    @mcp.custom_route(
        "/office-bridge/session", methods=["POST", "OPTIONS"], include_in_schema=True
    )
    async def pair_route(request: Request) -> Response:
        headers: dict[str, str] = {}
        try:
            headers = _cors_headers(request)
            if request.method == "OPTIONS":
                return Response(status_code=204, headers=headers)
            pairing_request = OfficePairSessionRequest.model_validate(
                await _read_json(request)
            )
            grant = await bridge.pair_session(pairing_request)
            return JSONResponse(
                grant.model_dump(mode="json"), status_code=201, headers=headers
            )
        except ValidationError:
            return _http_error(
                OfficeBridgeError("invalid_request", "Request body is invalid"),
                headers,
            )
        except OfficeBridgeError as exc:
            return _http_error(exc, headers)

    @mcp.custom_route(
        "/office-bridge/commands/poll",
        methods=["POST", "OPTIONS"],
        include_in_schema=True,
    )
    async def poll_route(request: Request) -> Response:
        headers: dict[str, str] = {}
        try:
            headers = _cors_headers(request)
            if request.method == "OPTIONS":
                return Response(status_code=204, headers=headers)
            token = _bearer_token(request)
            poll_request = OfficePollRequest.model_validate(await _read_json(request))
            command = await bridge.poll(token, poll_request.wait_seconds)
            response = OfficePollResponse(command=command)
            return JSONResponse(response.model_dump(mode="json"), headers=headers)
        except ValidationError:
            return _http_error(
                OfficeBridgeError("invalid_request", "Request body is invalid"),
                headers,
            )
        except OfficeBridgeError as exc:
            return _http_error(exc, headers)

    @mcp.custom_route(
        "/office-bridge/commands/result",
        methods=["POST", "OPTIONS"],
        include_in_schema=True,
    )
    async def result_route(request: Request) -> Response:
        headers: dict[str, str] = {}
        try:
            headers = _cors_headers(request)
            if request.method == "OPTIONS":
                return Response(status_code=204, headers=headers)
            token = _bearer_token(request)
            outcome = OUTCOME_ADAPTER.validate_python(await _read_json(request))
            receipt = await bridge.complete(token, outcome)
            return JSONResponse(receipt.model_dump(mode="json"), headers=headers)
        except ValidationError:
            return _http_error(
                OfficeBridgeError("invalid_request", "Request body is invalid"),
                headers,
            )
        except OfficeBridgeError as exc:
            return _http_error(exc, headers)

    @mcp.custom_route(
        "/office-bridge/session/disconnect",
        methods=["POST", "OPTIONS"],
        include_in_schema=True,
    )
    async def disconnect_route(request: Request) -> Response:
        headers: dict[str, str] = {}
        try:
            headers = _cors_headers(request)
            if request.method == "OPTIONS":
                return Response(status_code=204, headers=headers)
            token = _bearer_token(request)
            await bridge.disconnect(token)
            return JSONResponse({"disconnected": True}, headers=headers)
        except OfficeBridgeError as exc:
            return _http_error(exc, headers)
