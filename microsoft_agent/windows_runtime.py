"""Native, policy-enforcing runtime for the outbound Windows companion.

This module is deliberately a worker, not a web server.  A laptop polls an
authenticated control plane through an injected transport, executes only the
closed action types defined in :mod:`microsoft_agent.windows_companion`, and
acknowledges the result.  There is no command, shell, PowerShell, arbitrary
process, or arbitrary URL execution path.

The operating-system integrations are optional and injected.  The built-in
implementations use pywin32 (and, for toast notifications, win10toast) only on
Windows.  File access is rooted in an explicit logical-to-local binding and
rejects symlinks, junctions, and other reparse points before accessing data.
"""

from __future__ import annotations

import asyncio
import base64
import hashlib
import json
import ntpath
import os
import platform
import socket
import stat
import sys
from collections import OrderedDict
from collections.abc import Callable, Mapping
from datetime import UTC, datetime
from pathlib import Path, PureWindowsPath
from typing import Any, Protocol, TypeVar, runtime_checkable

from pydantic import BaseModel, ConfigDict, Field, field_validator

from microsoft_agent.windows_companion import (
    ActionPolicy,
    ClipboardReadTextAction,
    ClipboardWriteTextAction,
    CompanionActionKind,
    CompanionActionRequest,
    CompanionActionResult,
    CompanionActionStatus,
    CompanionDevice,
    DeviceIdentity,
    FileListAction,
    FileReadAction,
    FileWriteAction,
    NotificationShowAction,
    OfficeExportPdfAction,
    OfficeOpenDocumentAction,
    PowerAutomateDesktopRunAction,
    SystemInventoryAction,
    WindowsServiceStartAction,
    WindowsServiceStatusAction,
    WindowsServiceStopAction,
)

_REPARSE_POINT_ATTRIBUTE = 0x400
_T = TypeVar("_T")


class WindowsRuntimeLimits(BaseModel):
    """Local resource limits enforced independently of controller input."""

    model_config = ConfigDict(frozen=True)

    maximum_file_bytes: int = Field(default=10_485_760, ge=1, le=104_857_600)
    list_hash_maximum_bytes: int = Field(default=10_485_760, ge=0, le=104_857_600)
    maximum_pad_input_bytes: int = Field(default=262_144, ge=1, le=1_048_576)
    maximum_cached_results: int = Field(default=1024, ge=1, le=10_000)
    relay_batch_size: int = Field(default=10, ge=1, le=100)
    relay_wait_seconds: float = Field(default=30.0, ge=1, le=300)
    idle_delay_seconds: float = Field(default=1.0, ge=0.05, le=60)


class WindowsRuntimeError(RuntimeError):
    """Safe runtime failure with a stable machine-readable code."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        self.safe_message = message
        super().__init__(message)


class WindowsRuntimePolicyError(WindowsRuntimeError):
    """An action was rejected by the device-local policy."""

    def __init__(self, message: str, *, expired: bool = False) -> None:
        super().__init__("expired" if expired else "policy_denied", message)
        self.expired = expired


@runtime_checkable
class OfficeAutomation(Protocol):
    """Allowlisted Office automation operations."""

    def open_document(
        self, application: str, document_path: Path, mode: str
    ) -> Mapping[str, Any]:
        """Open one already-validated Office document."""

    def export_pdf(
        self, application: str, source_path: Path, output_path: Path
    ) -> Mapping[str, Any]:
        """Export one already-validated Office document as PDF."""


@runtime_checkable
class WindowsServiceManager(Protocol):
    """Narrow Windows service-control interface."""

    def status(self, service_name: str) -> str:
        """Return the status of a validated service name."""

    def start(self, service_name: str) -> str:
        """Start a validated service and return its new status."""

    def stop(self, service_name: str) -> str:
        """Stop a validated service and return its new status."""


@runtime_checkable
class ClipboardAdapter(Protocol):
    """Narrow text-only clipboard interface."""

    def read_text(self) -> str:
        """Read Unicode text from the interactive user's clipboard."""

    def write_text(self, value: str) -> None:
        """Replace the interactive user's clipboard with Unicode text."""


@runtime_checkable
class NotificationAdapter(Protocol):
    """Narrow local notification interface."""

    def show(self, title: str, message: str) -> None:
        """Show a bounded local notification."""


@runtime_checkable
class DesktopFlowExecutor(Protocol):
    """Injected Power Automate Desktop runner with no process primitive."""

    async def run_flow(
        self,
        flow_name: str,
        inputs: Mapping[str, Any],
        *,
        wait_for_completion: bool,
    ) -> Mapping[str, Any]:
        """Run one caller-validated, allowlisted desktop flow."""


class PyWin32OfficeAutomation:
    """Office COM automation backed by optional pywin32."""

    def __init__(self) -> None:
        _require_windows("Office COM automation")
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:  # pragma: no cover - Windows-only dependency
            raise ImportError(
                "Install pywin32 to enable Office COM automation"
            ) from exc
        self._pythoncom = pythoncom
        self._client = win32com.client

    def open_document(
        self, application: str, document_path: Path, mode: str
    ) -> Mapping[str, Any]:
        self._pythoncom.CoInitialize()
        try:
            if application == "word":
                app = self._client.DispatchEx("Word.Application")
                app.Visible = True
                app.Documents.Open(str(document_path), ReadOnly=mode == "view")
            elif application == "powerpoint":
                app = self._client.DispatchEx("PowerPoint.Application")
                app.Visible = True
                app.Presentations.Open(
                    str(document_path),
                    ReadOnly=mode == "view",
                    WithWindow=True,
                )
            elif application == "excel":
                app = self._client.DispatchEx("Excel.Application")
                app.Visible = True
                app.Workbooks.Open(str(document_path), ReadOnly=mode == "view")
            else:  # The action model already prevents this branch.
                raise WindowsRuntimeError(
                    "invalid_application", "Unsupported Office application"
                )
            return {"application": application, "mode": mode, "opened": True}
        finally:
            self._pythoncom.CoUninitialize()

    def export_pdf(
        self, application: str, source_path: Path, output_path: Path
    ) -> Mapping[str, Any]:
        self._pythoncom.CoInitialize()
        app: Any | None = None
        document: Any | None = None
        try:
            if application == "word":
                app = self._client.DispatchEx("Word.Application")
                app.Visible = False
                document = app.Documents.Open(str(source_path), ReadOnly=True)
                document.ExportAsFixedFormat(str(output_path), 17)
            elif application == "powerpoint":
                app = self._client.DispatchEx("PowerPoint.Application")
                app.Visible = False
                document = app.Presentations.Open(
                    str(source_path), ReadOnly=True, WithWindow=False
                )
                document.SaveAs(str(output_path), 32)
            elif application == "excel":
                app = self._client.DispatchEx("Excel.Application")
                app.Visible = False
                document = app.Workbooks.Open(str(source_path), ReadOnly=True)
                document.ExportAsFixedFormat(0, str(output_path))
            else:
                raise WindowsRuntimeError(
                    "invalid_application", "Unsupported Office application"
                )
            return {"application": application, "exported": True}
        finally:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:  # pragma: no cover - defensive COM cleanup
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:  # pragma: no cover - defensive COM cleanup
                    pass
            self._pythoncom.CoUninitialize()


class PyWin32ServiceManager:
    """Windows service manager backed by optional pywin32."""

    _STATE_NAMES = {
        1: "stopped",
        2: "start_pending",
        3: "stop_pending",
        4: "running",
        5: "continue_pending",
        6: "pause_pending",
        7: "paused",
    }

    def __init__(self) -> None:
        _require_windows("Windows service control")
        try:
            import win32serviceutil
        except ImportError as exc:  # pragma: no cover - Windows-only dependency
            raise ImportError("Install pywin32 to control Windows services") from exc
        self._service = win32serviceutil

    def status(self, service_name: str) -> str:
        state = int(self._service.QueryServiceStatus(service_name)[1])
        return self._STATE_NAMES.get(state, f"unknown_{state}")

    def start(self, service_name: str) -> str:
        self._service.StartService(service_name)
        return self.status(service_name)

    def stop(self, service_name: str) -> str:
        self._service.StopService(service_name)
        return self.status(service_name)


class PyWin32ClipboardAdapter:
    """Text-only clipboard adapter backed by optional pywin32."""

    def __init__(self) -> None:
        _require_windows("Windows clipboard access")
        try:
            import win32clipboard
            import win32con
        except ImportError as exc:  # pragma: no cover - Windows-only dependency
            raise ImportError(
                "Install pywin32 to access the Windows clipboard"
            ) from exc
        self._clipboard = win32clipboard
        self._unicode_format = win32con.CF_UNICODETEXT

    def read_text(self) -> str:
        self._clipboard.OpenClipboard()
        try:
            if not self._clipboard.IsClipboardFormatAvailable(self._unicode_format):
                return ""
            value = self._clipboard.GetClipboardData(self._unicode_format)
            return value if isinstance(value, str) else str(value)
        finally:
            self._clipboard.CloseClipboard()

    def write_text(self, value: str) -> None:
        self._clipboard.OpenClipboard()
        try:
            self._clipboard.EmptyClipboard()
            self._clipboard.SetClipboardText(value, self._unicode_format)
        finally:
            self._clipboard.CloseClipboard()


class WindowsToastNotificationAdapter:
    """Optional local toast notifications backed by win10toast."""

    def __init__(self) -> None:
        _require_windows("Windows notifications")
        try:
            from win10toast import ToastNotifier
        except ImportError as exc:  # pragma: no cover - optional dependency
            raise ImportError(
                "Install win10toast to show Windows notifications"
            ) from exc
        self._notifier = ToastNotifier()

    def show(self, title: str, message: str) -> None:
        self._notifier.show_toast(title, message, threaded=False, duration=5)


class _RootBinding:
    def __init__(self, logical_root: str, physical_root: Path) -> None:
        self.logical_root = _normalize_windows_path(logical_root)
        self.physical_root = physical_root.absolute()
        _reject_reparse(self.physical_root)
        if not self.physical_root.is_dir():
            raise ValueError(
                f"Bound file root for {logical_root!r} must be an existing directory"
            )
        self.resolved_root = self.physical_root.resolve(strict=True)
        _ensure_contained(self.resolved_root, self.physical_root.resolve(strict=True))
        root_stat = os.lstat(self.physical_root)
        self.root_identity = (root_stat.st_dev, root_stat.st_ino)

    def validate_root(self) -> None:
        _reject_reparse(self.physical_root)
        current = os.lstat(self.physical_root)
        if (current.st_dev, current.st_ino) != self.root_identity:
            raise WindowsRuntimePolicyError("An allowlisted file root was replaced")
        if self.physical_root.resolve(strict=True) != self.resolved_root:
            raise WindowsRuntimePolicyError("An allowlisted file root was redirected")


class _GuardedPath(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True, frozen=True)

    logical_path: str
    physical_path: Path
    binding: Any


class _FileGuard:
    def __init__(
        self,
        device: CompanionDevice,
        bindings: Mapping[str, str | os.PathLike[str]] | None,
    ) -> None:
        supplied = {
            _normalize_windows_path(key): Path(value)
            for key, value in (bindings or {}).items()
        }
        allowed = {_normalize_windows_path(root) for root in device.allowed_file_roots}
        extras = set(supplied) - allowed
        if extras:
            raise ValueError(
                "File-root bindings contain roots not allowed by the device"
            )
        if os.name != "nt" and allowed - set(supplied):
            raise ValueError(
                "Non-Windows runtimes require a local binding for every allowed file root"
            )
        self._bindings: list[_RootBinding] = []
        for root in allowed:
            physical = supplied.get(root, Path(root))
            self._bindings.append(_RootBinding(root, physical))
        self._bindings.sort(key=lambda item: len(item.logical_root), reverse=True)

    def read_path(self, value: str, *, require_directory: bool = False) -> _GuardedPath:
        guarded = self._translate(value)
        self._validate_chain(guarded, require_target=True)
        if require_directory and not guarded.physical_path.is_dir():
            raise WindowsRuntimePolicyError("The requested path is not a directory")
        return guarded

    def write_path(self, value: str) -> _GuardedPath:
        guarded = self._translate(value)
        self._validate_chain(guarded, require_target=False)
        parent = guarded.physical_path.parent
        if not parent.is_dir():
            raise WindowsRuntimePolicyError("The destination parent does not exist")
        _reject_reparse(parent)
        _ensure_contained(parent.resolve(strict=True), guarded.binding.resolved_root)
        if guarded.physical_path.exists():
            _reject_reparse(guarded.physical_path)
        return guarded

    def logical_child(self, guarded: _GuardedPath, child: Path) -> str:
        relative = child.relative_to(guarded.physical_path)
        return ntpath.normpath(ntpath.join(guarded.logical_path, *relative.parts))

    def validate_existing(self, guarded: _GuardedPath) -> None:
        self._validate_chain(guarded, require_target=True)

    def _translate(self, value: str) -> _GuardedPath:
        candidate_raw = ntpath.normpath(value.strip())
        candidate = _normalize_windows_path(candidate_raw)
        if not ntpath.isabs(candidate):
            raise WindowsRuntimePolicyError("File paths must be absolute Windows paths")
        if candidate.startswith(("\\\\?\\", "\\\\.\\", "\\??\\")):
            raise WindowsRuntimePolicyError(
                "Windows device namespace paths are forbidden"
            )
        for binding in self._bindings:
            try:
                if (
                    ntpath.commonpath((candidate, binding.logical_root))
                    != binding.logical_root
                ):
                    continue
            except ValueError:
                continue
            relative = ntpath.relpath(candidate_raw, binding.logical_root)
            parts = () if relative == "." else PureWindowsPath(relative).parts
            if any(part in {"", ".", ".."} or ":" in part for part in parts):
                raise WindowsRuntimePolicyError("Unsafe Windows path component")
            physical = binding.physical_root.joinpath(*parts)
            return _GuardedPath(
                logical_path=candidate_raw,
                physical_path=physical,
                binding=binding,
            )
        raise WindowsRuntimePolicyError("File path is outside the device allowlist")

    @staticmethod
    def _validate_chain(guarded: _GuardedPath, *, require_target: bool) -> None:
        binding: _RootBinding = guarded.binding
        binding.validate_root()
        candidate = guarded.physical_path
        current = binding.physical_root
        try:
            relative_parts = candidate.relative_to(binding.physical_root).parts
        except ValueError as exc:
            raise WindowsRuntimePolicyError("File path escaped its bound root") from exc
        for index, part in enumerate(relative_parts):
            current = current / part
            is_target = index == len(relative_parts) - 1
            if not current.exists() and not current.is_symlink():
                if require_target or not is_target:
                    raise WindowsRuntimePolicyError("The requested path does not exist")
                break
            _reject_reparse(current)
        resolved = candidate.resolve(strict=require_target)
        _ensure_contained(resolved, binding.resolved_root)


class WindowsActionExecutor:
    """Execute typed actions after enforcing device-local policy and limits."""

    def __init__(
        self,
        device: CompanionDevice,
        *,
        file_root_bindings: Mapping[str, str | os.PathLike[str]] | None = None,
        action_policies: Mapping[CompanionActionKind, ActionPolicy] | None = None,
        limits: WindowsRuntimeLimits | None = None,
        office: OfficeAutomation | None = None,
        services: WindowsServiceManager | None = None,
        clipboard: ClipboardAdapter | None = None,
        notifications: NotificationAdapter | None = None,
        desktop_flows: DesktopFlowExecutor | None = None,
        clock: Callable[[], datetime] | None = None,
    ) -> None:
        self.device = device
        self.limits = limits or WindowsRuntimeLimits()
        self._policies = dict(action_policies or _runtime_action_policies())
        missing = set(device.allowed_actions) - set(self._policies)
        if missing:
            raise ValueError(f"Local action policies are missing: {sorted(missing)}")
        self._files = _FileGuard(device, file_root_bindings)
        self._office = office
        self._services = services
        self._clipboard = clipboard
        self._notifications = notifications
        self._desktop_flows = desktop_flows
        self._clock = clock or (lambda: datetime.now(UTC))
        self._cache: OrderedDict[str, tuple[str, CompanionActionResult]] = OrderedDict()
        self._inflight: dict[
            str, tuple[str, asyncio.Future[CompanionActionResult]]
        ] = {}
        self._cache_lock = asyncio.Lock()

    @property
    def local_policies(self) -> Mapping[CompanionActionKind, ActionPolicy]:
        """Return a copy of the immutable local policy decision table."""

        return dict(self._policies)

    async def execute(self, request: CompanionActionRequest) -> CompanionActionResult:
        """Execute one action with idempotent result replay."""

        fingerprint = _request_fingerprint(request)
        owner = False
        async with self._cache_lock:
            cached = self._cache.get(request.idempotency_key)
            if cached is not None:
                known_fingerprint, result = cached
                if known_fingerprint != fingerprint:
                    return self._rejected(
                        request,
                        "idempotency_conflict",
                        "The idempotency key was already used for another request",
                    )
                self._cache.move_to_end(request.idempotency_key)
                return result
            pending = self._inflight.get(request.idempotency_key)
            if pending is not None:
                known_fingerprint, future = pending
                if known_fingerprint != fingerprint:
                    return self._rejected(
                        request,
                        "idempotency_conflict",
                        "The idempotency key is in use by another request",
                    )
            else:
                future = asyncio.get_running_loop().create_future()
                self._inflight[request.idempotency_key] = (fingerprint, future)
                owner = True
        if not owner:
            return await asyncio.shield(future)

        try:
            result = await self._execute_uncached(request)
        except BaseException as exc:
            async with self._cache_lock:
                self._inflight.pop(request.idempotency_key, None)
                if not future.done():
                    future.set_exception(exc)
                    future.exception()
            raise

        async with self._cache_lock:
            self._inflight.pop(request.idempotency_key, None)
            self._cache[request.idempotency_key] = (fingerprint, result)
            self._cache.move_to_end(request.idempotency_key)
            while len(self._cache) > self.limits.maximum_cached_results:
                self._cache.popitem(last=False)
            if not future.done():
                future.set_result(result)
        return result

    def authorize(self, request: CompanionActionRequest) -> ActionPolicy:
        """Validate device, action allowlists, expiry, and confirmation evidence."""

        if not self.device.enabled:
            raise WindowsRuntimePolicyError("This device is disabled")
        kind = CompanionActionKind(request.action.kind)
        if kind not in self.device.allowed_actions:
            raise WindowsRuntimePolicyError(
                f"Action {kind.value!r} is not allowed on this device"
            )
        policy = self._policies.get(kind)
        if policy is None:
            raise WindowsRuntimePolicyError("No local policy exists for this action")
        now = self._clock()
        if now.utcoffset() is None:
            raise RuntimeError(
                "The runtime clock must return a timezone-aware datetime"
            )
        if request.expires_at <= now:
            raise WindowsRuntimePolicyError(
                "The action request has expired", expired=True
            )

        action = request.action
        if isinstance(action, (FileListAction, FileReadAction)):
            self._files.read_path(
                action.path, require_directory=isinstance(action, FileListAction)
            )
        elif isinstance(action, FileWriteAction):
            self._files.write_path(action.path)
        elif isinstance(action, OfficeOpenDocumentAction):
            self._files.read_path(action.document_path)
        elif isinstance(action, OfficeExportPdfAction):
            self._files.read_path(action.source_path)
            self._files.write_path(action.output_path)
        elif isinstance(action, PowerAutomateDesktopRunAction):
            if action.flow_name.casefold() not in {
                name.casefold() for name in self.device.allowed_desktop_flows
            }:
                raise WindowsRuntimePolicyError(
                    "Power Automate Desktop flow is not allowlisted"
                )
        elif isinstance(
            action,
            (
                WindowsServiceStatusAction,
                WindowsServiceStartAction,
                WindowsServiceStopAction,
            ),
        ) and action.service_name.casefold() not in {
            name.casefold() for name in self.device.allowed_services
        }:
            raise WindowsRuntimePolicyError("Windows service is not allowlisted")

        if policy.requires_confirmation:
            evidence = request.confirmation
            if evidence is None:
                raise WindowsRuntimePolicyError("This action requires confirmation")
            if evidence.action_kind is not kind:
                raise WindowsRuntimePolicyError(
                    "Confirmation was granted for a different action kind"
                )
            if evidence.confirmed_at > now:
                raise WindowsRuntimePolicyError("Confirmation is dated in the future")
            if evidence.expires_at <= now:
                raise WindowsRuntimePolicyError("Confirmation has expired")
        return policy

    async def _execute_uncached(
        self, request: CompanionActionRequest
    ) -> CompanionActionResult:
        started = self._clock()
        try:
            self.authorize(request)
            output = await self._dispatch(request)
        except WindowsRuntimePolicyError as exc:
            return self._rejected(
                request,
                exc.code,
                exc.safe_message,
                status=(
                    CompanionActionStatus.EXPIRED
                    if exc.expired
                    else CompanionActionStatus.REJECTED
                ),
                started_at=started,
            )
        except WindowsRuntimeError as exc:
            return self._failed(request, exc.code, exc.safe_message, started)
        except Exception as exc:
            return self._failed(
                request,
                "execution_failed",
                f"Action execution failed: {type(exc).__name__}",
                started,
            )
        return CompanionActionResult(
            action_id=request.action_id,
            device_id=self.device.identity.device_id,
            status=CompanionActionStatus.SUCCEEDED,
            started_at=started,
            completed_at=self._clock(),
            output=dict(output),
        )

    async def _dispatch(self, request: CompanionActionRequest) -> Mapping[str, Any]:
        action = request.action
        if isinstance(action, SystemInventoryAction):
            return await asyncio.to_thread(self._inventory, action)
        if isinstance(action, FileListAction):
            return await asyncio.to_thread(self._list_files, action)
        if isinstance(action, FileReadAction):
            return await asyncio.to_thread(self._read_file, action)
        if isinstance(action, FileWriteAction):
            return await asyncio.to_thread(self._write_file, action)
        if isinstance(action, OfficeOpenDocumentAction):
            return await self._open_office(action)
        if isinstance(action, OfficeExportPdfAction):
            return await self._export_office(action)
        if isinstance(action, PowerAutomateDesktopRunAction):
            return await self._run_desktop_flow(action)
        if isinstance(action, WindowsServiceStatusAction):
            return await self._service_action("status", action.service_name)
        if isinstance(action, WindowsServiceStartAction):
            return await self._service_action("start", action.service_name)
        if isinstance(action, WindowsServiceStopAction):
            return await self._service_action("stop", action.service_name)
        if isinstance(action, NotificationShowAction):
            notification_adapter = self._require_adapter(
                self._notifications, "notifications"
            )
            await asyncio.to_thread(
                notification_adapter.show, action.title, action.message
            )
            return {"shown": True}
        if isinstance(action, ClipboardReadTextAction):
            clipboard_adapter = self._require_adapter(self._clipboard, "clipboard")
            value = await asyncio.to_thread(clipboard_adapter.read_text)
            truncated = len(value) > action.max_characters
            value = value[: action.max_characters]
            return {
                "text": value,
                "characters": len(value),
                "truncated": truncated,
                "sha256": hashlib.sha256(value.encode("utf-8")).hexdigest(),
            }
        if isinstance(action, ClipboardWriteTextAction):
            clipboard_adapter = self._require_adapter(self._clipboard, "clipboard")
            await asyncio.to_thread(clipboard_adapter.write_text, action.text)
            return {
                "characters": len(action.text),
                "sha256": hashlib.sha256(action.text.encode("utf-8")).hexdigest(),
            }
        raise WindowsRuntimeError("unsupported_action", "Unsupported companion action")

    def _inventory(self, action: SystemInventoryAction) -> Mapping[str, Any]:
        uname = platform.uname()
        output: dict[str, Any] = {
            "hostname": socket.gethostname(),
            "platform": platform.platform(),
            "system": uname.system,
            "release": uname.release,
            "version": uname.version,
            "machine": uname.machine,
            "processor": uname.processor,
            "python_version": platform.python_version(),
        }
        if action.include_network_adapters:
            addresses: set[str] = set()
            try:
                for item in socket.getaddrinfo(socket.gethostname(), None):
                    address = item[4][0]
                    if isinstance(address, str):
                        addresses.add(address.split("%", 1)[0])
            except OSError:
                pass
            output["network_addresses"] = sorted(addresses)[:128]
        if action.include_software:
            output["installed_software"] = _installed_windows_software()
        return output

    def _list_files(self, action: FileListAction) -> Mapping[str, Any]:
        guarded = self._files.read_path(action.path, require_directory=True)
        entries: list[dict[str, Any]] = []
        truncated = False
        pending = [guarded.physical_path]
        while pending:
            directory = pending.pop()
            try:
                children = sorted(
                    os.scandir(directory), key=lambda item: item.name.casefold()
                )
            except OSError as exc:
                raise WindowsRuntimeError(
                    "file_access_failed", "The directory could not be listed"
                ) from exc
            for entry in children:
                if len(entries) >= action.max_entries:
                    truncated = True
                    pending.clear()
                    break
                path = Path(entry.path)
                info = entry.stat(follow_symlinks=False)
                reparse = _stat_is_reparse(info)
                is_directory = stat.S_ISDIR(info.st_mode) and not reparse
                is_file = stat.S_ISREG(info.st_mode) and not reparse
                digest: str | None = None
                if is_file and info.st_size <= self.limits.list_hash_maximum_bytes:
                    digest = _hash_file_no_follow(path, self.limits.maximum_file_bytes)
                entries.append(
                    {
                        "path": self._files.logical_child(guarded, path),
                        "name": entry.name,
                        "type": (
                            "reparse_point"
                            if reparse
                            else "directory"
                            if is_directory
                            else "file"
                            if is_file
                            else "other"
                        ),
                        "size_bytes": info.st_size if is_file else None,
                        "modified_at": datetime.fromtimestamp(
                            info.st_mtime, tz=UTC
                        ).isoformat(),
                        "sha256": digest,
                    }
                )
                if action.recursive and is_directory:
                    pending.append(path)
        return {"entries": entries, "count": len(entries), "truncated": truncated}

    def _read_file(self, action: FileReadAction) -> Mapping[str, Any]:
        guarded = self._files.read_path(action.path)
        maximum = min(action.max_bytes, self.limits.maximum_file_bytes)
        data = _read_file_no_follow(guarded.physical_path, maximum)
        self._files.validate_existing(guarded)
        return {
            "path": guarded.logical_path,
            "content_base64": base64.b64encode(data).decode("ascii"),
            "size_bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
        }

    def _write_file(self, action: FileWriteAction) -> Mapping[str, Any]:
        guarded = self._files.write_path(action.path)
        data = base64.b64decode(action.content_base64, validate=True)
        if len(data) > self.limits.maximum_file_bytes:
            raise WindowsRuntimePolicyError("File content exceeds the local byte limit")
        digest = hashlib.sha256(data).hexdigest()
        if (
            action.expected_sha256
            and digest.casefold() != action.expected_sha256.casefold()
        ):
            raise WindowsRuntimePolicyError(
                "File content SHA-256 did not match expectation"
            )
        target = guarded.physical_path
        if target.exists() and not action.overwrite:
            raise WindowsRuntimePolicyError(
                "Destination exists and overwrite is disabled"
            )
        _atomic_write(target, data, overwrite=action.overwrite)
        self._files.validate_existing(guarded)
        written_digest = _hash_file_no_follow(target, self.limits.maximum_file_bytes)
        if written_digest != digest:
            raise WindowsRuntimeError(
                "integrity_check_failed", "Written file failed its integrity check"
            )
        return {
            "path": guarded.logical_path,
            "size_bytes": len(data),
            "sha256": digest,
            "overwritten": action.overwrite,
        }

    async def _open_office(self, action: OfficeOpenDocumentAction) -> Mapping[str, Any]:
        adapter = self._require_adapter(self._office, "Office automation")
        guarded = self._files.read_path(action.document_path)
        result = await asyncio.to_thread(
            adapter.open_document,
            action.application,
            guarded.physical_path,
            action.mode,
        )
        return {**dict(result), "document_path": guarded.logical_path}

    async def _export_office(self, action: OfficeExportPdfAction) -> Mapping[str, Any]:
        adapter = self._require_adapter(self._office, "Office automation")
        source = self._files.read_path(action.source_path)
        output = self._files.write_path(action.output_path)
        if output.physical_path.exists() and not action.overwrite:
            raise WindowsRuntimePolicyError(
                "PDF destination exists and overwrite is disabled"
            )
        result = await asyncio.to_thread(
            adapter.export_pdf,
            action.application,
            source.physical_path,
            output.physical_path,
        )
        self._files.validate_existing(output)
        if not output.physical_path.is_file():
            raise WindowsRuntimeError(
                "office_export_failed", "Office did not produce the requested PDF"
            )
        return {
            **dict(result),
            "source_path": source.logical_path,
            "output_path": output.logical_path,
            "size_bytes": output.physical_path.stat().st_size,
            "sha256": _hash_file_no_follow(
                output.physical_path, self.limits.maximum_file_bytes
            ),
        }

    async def _run_desktop_flow(
        self, action: PowerAutomateDesktopRunAction
    ) -> Mapping[str, Any]:
        executor = self._require_adapter(self._desktop_flows, "desktop flow execution")
        try:
            encoded = json.dumps(
                action.inputs, ensure_ascii=False, separators=(",", ":"), sort_keys=True
            ).encode("utf-8")
        except (TypeError, ValueError) as exc:
            raise WindowsRuntimePolicyError(
                "Desktop flow inputs must be JSON serializable"
            ) from exc
        if len(encoded) > self.limits.maximum_pad_input_bytes:
            raise WindowsRuntimePolicyError(
                "Desktop flow inputs exceed the local limit"
            )
        result = await executor.run_flow(
            action.flow_name,
            action.inputs,
            wait_for_completion=action.wait_for_completion,
        )
        try:
            json.dumps(result)
        except (TypeError, ValueError) as exc:
            raise WindowsRuntimeError(
                "invalid_desktop_flow_result",
                "Desktop flow executor returned a non-JSON result",
            ) from exc
        return dict(result)

    async def _service_action(
        self, operation: str, service_name: str
    ) -> Mapping[str, Any]:
        adapter = self._require_adapter(self._services, "Windows service control")
        method = getattr(adapter, operation)
        status = await asyncio.to_thread(method, service_name)
        return {"service_name": service_name, "status": status, "operation": operation}

    @staticmethod
    def _require_adapter(value: _T | None, capability: str) -> _T:
        if value is None:
            raise WindowsRuntimeError(
                "dependency_unavailable",
                f"{capability} is not configured on this device",
            )
        return value

    def _rejected(
        self,
        request: CompanionActionRequest,
        code: str,
        message: str,
        *,
        status: CompanionActionStatus = CompanionActionStatus.REJECTED,
        started_at: datetime | None = None,
    ) -> CompanionActionResult:
        from microsoft_agent.windows_companion import CompanionActionFailure

        return CompanionActionResult(
            action_id=request.action_id,
            device_id=self.device.identity.device_id,
            status=status,
            started_at=started_at,
            completed_at=self._clock(),
            error=CompanionActionFailure(code=code, message=message, retryable=False),
        )

    def _failed(
        self,
        request: CompanionActionRequest,
        code: str,
        message: str,
        started_at: datetime,
    ) -> CompanionActionResult:
        from microsoft_agent.windows_companion import CompanionActionFailure

        return CompanionActionResult(
            action_id=request.action_id,
            device_id=self.device.identity.device_id,
            status=CompanionActionStatus.FAILED,
            started_at=started_at,
            completed_at=self._clock(),
            error=CompanionActionFailure(code=code, message=message, retryable=False),
        )


class RelayActionDelivery(BaseModel):
    """One authenticated control-plane delivery to the outbound worker."""

    model_config = ConfigDict(frozen=True)

    delivery_id: str = Field(min_length=1, max_length=128)
    device_id: str = Field(min_length=1, max_length=128)
    expected_device_identity: DeviceIdentity
    request: CompanionActionRequest
    policy: ActionPolicy

    @field_validator("delivery_id")
    @classmethod
    def validate_delivery_id(cls, value: str) -> str:
        if any(char.isspace() or not 33 <= ord(char) <= 126 for char in value):
            raise ValueError("delivery_id must use visible ASCII characters")
        return value


class RelayPollBatch(BaseModel):
    """Bounded set of action deliveries received through an outbound poll."""

    model_config = ConfigDict(frozen=True)

    cursor: str | None = Field(default=None, max_length=1024)
    deliveries: tuple[RelayActionDelivery, ...] = ()


@runtime_checkable
class OutboundRelayTransport(Protocol):
    """Injected authenticated long-poll and acknowledgement transport."""

    async def poll(
        self,
        identity: DeviceIdentity,
        *,
        cursor: str | None,
        maximum_actions: int,
        wait_seconds: float,
    ) -> RelayPollBatch:
        """Poll the control plane over an authenticated outbound connection."""

    async def acknowledge(
        self,
        identity: DeviceIdentity,
        delivery_id: str,
        result: CompanionActionResult,
    ) -> None:
        """Acknowledge a delivery and its locally-produced result."""


class OutboundRelayWorker:
    """Poll, execute, and acknowledge actions without opening an inbound port."""

    def __init__(
        self,
        transport: OutboundRelayTransport,
        executor: WindowsActionExecutor,
        *,
        limits: WindowsRuntimeLimits | None = None,
    ) -> None:
        self._transport = transport
        self._executor = executor
        self._limits = limits or executor.limits
        self._cursor: str | None = None

    @property
    def cursor(self) -> str | None:
        """Return the last fully acknowledged relay cursor."""

        return self._cursor

    async def run_once(self) -> int:
        """Process one bounded long-poll batch and return its delivery count."""

        identity = self._executor.device.identity
        batch = await self._transport.poll(
            identity,
            cursor=self._cursor,
            maximum_actions=self._limits.relay_batch_size,
            wait_seconds=self._limits.relay_wait_seconds,
        )
        if len(batch.deliveries) > self._limits.relay_batch_size:
            raise WindowsRuntimeError(
                "invalid_relay_batch", "Relay returned more actions than requested"
            )
        for delivery in batch.deliveries:
            result = await self._execute_delivery(delivery)
            await self._transport.acknowledge(identity, delivery.delivery_id, result)
        self._cursor = batch.cursor
        return len(batch.deliveries)

    async def run(self, stop_event: asyncio.Event) -> None:
        """Run outbound polling until ``stop_event`` is set."""

        while not stop_event.is_set():
            count = await self.run_once()
            if count:
                continue
            try:
                await asyncio.wait_for(
                    stop_event.wait(), timeout=self._limits.idle_delay_seconds
                )
            except TimeoutError:
                pass

    async def _execute_delivery(
        self, delivery: RelayActionDelivery
    ) -> CompanionActionResult:
        identity = self._executor.device.identity
        if delivery.device_id != identity.device_id:
            return self._executor._rejected(  # noqa: SLF001 - worker owns executor
                delivery.request,
                "identity_mismatch",
                "Relay delivery targeted a different device",
            )
        if delivery.expected_device_identity != identity:
            return self._executor._rejected(  # noqa: SLF001 - worker owns executor
                delivery.request,
                "identity_mismatch",
                "Relay delivery identity did not match local identity",
            )
        kind = CompanionActionKind(delivery.request.action.kind)
        local_policy = self._executor.local_policies.get(kind)
        if local_policy is None or delivery.policy != local_policy:
            return self._executor._rejected(  # noqa: SLF001 - worker owns executor
                delivery.request,
                "policy_mismatch",
                "Controller policy did not match the device-local policy",
            )
        return await self._executor.execute(delivery.request)


def _runtime_action_policies() -> dict[CompanionActionKind, ActionPolicy]:
    # Keep a device-local copy rather than trusting policy supplied by a relay.
    from microsoft_agent.windows_companion import (
        ConfirmationRequirement,
    )

    read = ActionPolicy(
        confirmation=ConfirmationRequirement.NONE,
        rationale="Read-only device metadata",
    )
    sensitive_read = ActionPolicy(
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
        CompanionActionKind.FILE_LIST: sensitive_read,
        CompanionActionKind.FILE_READ: sensitive_read,
        CompanionActionKind.FILE_WRITE: change,
        CompanionActionKind.OFFICE_OPEN_DOCUMENT: change,
        CompanionActionKind.OFFICE_EXPORT_PDF: change,
        CompanionActionKind.POWER_AUTOMATE_DESKTOP_RUN: change,
        CompanionActionKind.WINDOWS_SERVICE_STATUS: read,
        CompanionActionKind.WINDOWS_SERVICE_START: change,
        CompanionActionKind.WINDOWS_SERVICE_STOP: change,
        CompanionActionKind.NOTIFICATION_SHOW: change,
        CompanionActionKind.CLIPBOARD_READ_TEXT: sensitive_read,
        CompanionActionKind.CLIPBOARD_WRITE_TEXT: change,
    }


def _normalize_windows_path(value: str) -> str:
    if not value or "\x00" in value:
        raise ValueError("Windows path must be non-empty and contain no NUL")
    return ntpath.normcase(ntpath.normpath(value.strip()))


def _stat_is_reparse(value: os.stat_result) -> bool:
    attributes = int(getattr(value, "st_file_attributes", 0))
    return stat.S_ISLNK(value.st_mode) or bool(attributes & _REPARSE_POINT_ATTRIBUTE)


def _reject_reparse(path: Path) -> None:
    try:
        value = os.lstat(path)
    except FileNotFoundError:
        return
    if _stat_is_reparse(value):
        raise WindowsRuntimePolicyError(
            "Symlinks and Windows reparse points are forbidden"
        )


def _ensure_contained(candidate: Path, root: Path) -> None:
    try:
        if os.path.commonpath((str(candidate), str(root))) == str(root):
            return
    except ValueError:
        pass
    raise WindowsRuntimePolicyError("File path escaped its configured root")


def _read_file_no_follow(path: Path, maximum_bytes: int) -> bytes:
    _reject_reparse(path)
    flags = os.O_RDONLY | getattr(os, "O_BINARY", 0)
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    try:
        descriptor = os.open(path, flags)
    except OSError as exc:
        raise WindowsRuntimeError(
            "file_access_failed", "The file could not be opened"
        ) from exc
    try:
        info = os.fstat(descriptor)
        if not stat.S_ISREG(info.st_mode):
            raise WindowsRuntimePolicyError("The requested path is not a regular file")
        if info.st_size > maximum_bytes:
            raise WindowsRuntimePolicyError("The file exceeds the requested byte limit")
        data = bytearray()
        while len(data) <= maximum_bytes:
            chunk = os.read(descriptor, min(1024 * 1024, maximum_bytes + 1 - len(data)))
            if not chunk:
                break
            data.extend(chunk)
        if len(data) > maximum_bytes:
            raise WindowsRuntimePolicyError("The file exceeds the requested byte limit")
        return bytes(data)
    finally:
        os.close(descriptor)


def _hash_file_no_follow(path: Path, maximum_bytes: int) -> str:
    return hashlib.sha256(_read_file_no_follow(path, maximum_bytes)).hexdigest()


def _atomic_write(path: Path, data: bytes, *, overwrite: bool) -> None:
    parent = path.parent
    _reject_reparse(parent)
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0)
    temporary: Path | None = None
    descriptor: int | None = None
    for nonce in range(100):
        candidate = parent / f".microsoft-agent-{os.getpid()}-{nonce}.tmp"
        try:
            descriptor = os.open(candidate, flags, 0o600)
            temporary = candidate
            break
        except FileExistsError:
            continue
    if descriptor is None or temporary is None:
        raise WindowsRuntimeError(
            "file_write_failed", "Could not allocate a temporary destination"
        )
    try:
        view = memoryview(data)
        while view:
            written = os.write(descriptor, view)
            view = view[written:]
        os.fsync(descriptor)
        os.close(descriptor)
        descriptor = None
        if overwrite:
            _reject_reparse(path)
            os.replace(temporary, path)
        else:
            try:
                os.link(temporary, path, follow_symlinks=False)
            except FileExistsError as exc:
                raise WindowsRuntimePolicyError(
                    "Destination exists and overwrite is disabled"
                ) from exc
            os.unlink(temporary)
            temporary = None
    except WindowsRuntimeError:
        raise
    except OSError as exc:
        raise WindowsRuntimeError(
            "file_write_failed", "The file could not be written"
        ) from exc
    finally:
        if descriptor is not None:
            os.close(descriptor)
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass


def _request_fingerprint(request: CompanionActionRequest) -> str:
    payload = json.dumps(
        request.model_dump(mode="json"),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _require_windows(capability: str) -> None:
    if sys.platform != "win32":
        raise RuntimeError(f"{capability} is available only on Windows")


def _installed_windows_software() -> list[dict[str, str]]:
    if sys.platform != "win32":
        return []
    try:
        import winreg
    except ImportError:  # pragma: no cover - part of CPython on Windows
        return []
    roots = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    )
    found: dict[tuple[str, str], dict[str, str]] = {}
    for root_name in roots:
        try:
            root = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, root_name)
        except OSError:
            continue
        with root:
            for index in range(500):
                try:
                    child_name = winreg.EnumKey(root, index)
                    child = winreg.OpenKey(root, child_name)
                except OSError:
                    break
                with child:
                    try:
                        name = str(winreg.QueryValueEx(child, "DisplayName")[0]).strip()
                    except OSError:
                        continue
                    try:
                        version = str(
                            winreg.QueryValueEx(child, "DisplayVersion")[0]
                        ).strip()
                    except OSError:
                        version = ""
                    if name:
                        found[(name.casefold(), version)] = {
                            "name": name[:512],
                            "version": version[:128],
                        }
    return sorted(found.values(), key=lambda item: item["name"].casefold())[:500]


__all__ = [
    "ClipboardAdapter",
    "DesktopFlowExecutor",
    "NotificationAdapter",
    "OfficeAutomation",
    "OutboundRelayTransport",
    "OutboundRelayWorker",
    "PyWin32ClipboardAdapter",
    "PyWin32OfficeAutomation",
    "PyWin32ServiceManager",
    "RelayActionDelivery",
    "RelayPollBatch",
    "WindowsActionExecutor",
    "WindowsRuntimeError",
    "WindowsRuntimeLimits",
    "WindowsRuntimePolicyError",
    "WindowsServiceManager",
    "WindowsToastNotificationAdapter",
]
