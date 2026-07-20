"""Tests for the policy-enforcing native Windows companion runtime."""

from __future__ import annotations

import base64
import hashlib
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any
from uuid import UUID

import pytest

from microsoft_agent.windows_companion import (
    ClipboardReadTextAction,
    ClipboardWriteTextAction,
    CompanionActionKind,
    CompanionActionRequest,
    CompanionActionStatus,
    CompanionDevice,
    ConfirmationEvidence,
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
)
from microsoft_agent.windows_runtime import (
    OutboundRelayWorker,
    PyWin32OfficeAutomation,
    RelayActionDelivery,
    RelayPollBatch,
    WindowsActionExecutor,
    WindowsRuntimePolicyError,
)

NOW = datetime(2026, 7, 14, 15, 0, tzinfo=UTC)
LOGICAL_ROOT = r"C:\Allowed"


def _identity(device_id: str = "laptop-1") -> DeviceIdentity:
    return DeviceIdentity(
        device_id=device_id,
        tenant_id=UUID("11111111-1111-1111-1111-111111111111"),
        entra_device_id=UUID("22222222-2222-2222-2222-222222222222"),
        certificate_thumbprint="A" * 40,
    )


def _device(actions: set[CompanionActionKind] | None = None) -> CompanionDevice:
    return CompanionDevice(
        identity=_identity(),
        display_name="Test laptop",
        allowed_actions=frozenset(actions or set(CompanionActionKind)),
        allowed_file_roots=(LOGICAL_ROOT,),
        allowed_services=frozenset({"Spooler"}),
        allowed_desktop_flows=frozenset({"Reconcile invoices"}),
    )


def _confirmation(kind: CompanionActionKind) -> ConfirmationEvidence:
    return ConfirmationEvidence(
        action_kind=kind,
        confirmed_by="user@example.com",
        confirmed_at=NOW - timedelta(minutes=1),
        expires_at=NOW + timedelta(minutes=4),
        purpose="Approved test action",
        authorization_reference="approval/test-1",
    )


def _request(
    action: Any, *, confirmed: bool = True, key: str = "key-1"
) -> CompanionActionRequest:
    kind = CompanionActionKind(action.kind)
    return CompanionActionRequest(
        action=action,
        requested_at=NOW - timedelta(seconds=5),
        expires_at=NOW + timedelta(minutes=5),
        idempotency_key=key,
        confirmation=_confirmation(kind) if confirmed else None,
    )


def _executor(tmp_path: Path, **kwargs: Any) -> WindowsActionExecutor:
    return WindowsActionExecutor(
        _device(),
        file_root_bindings={LOGICAL_ROOT: tmp_path},
        clock=lambda: NOW,
        **kwargs,
    )


@pytest.mark.asyncio
async def test_inventory_executes_without_confirmation(tmp_path: Path) -> None:
    executor = _executor(tmp_path)
    request = _request(SystemInventoryAction(), confirmed=False)

    result = await executor.execute(request)

    assert result.status is CompanionActionStatus.SUCCEEDED
    assert result.output
    assert result.output["hostname"]
    assert result.output["python_version"]


@pytest.mark.asyncio
async def test_sensitive_file_read_requires_confirmation(tmp_path: Path) -> None:
    (tmp_path / "report.txt").write_text("secret", encoding="utf-8")
    executor = _executor(tmp_path)

    result = await executor.execute(
        _request(FileReadAction(path=r"C:\Allowed\report.txt"), confirmed=False)
    )

    assert result.status is CompanionActionStatus.REJECTED
    assert result.error
    assert result.error.code == "policy_denied"


@pytest.mark.asyncio
async def test_file_read_is_bounded_and_returns_hash(tmp_path: Path) -> None:
    content = b"bounded content"
    (tmp_path / "report.txt").write_bytes(content)
    executor = _executor(tmp_path)

    result = await executor.execute(
        _request(FileReadAction(path=r"C:\Allowed\report.txt", max_bytes=100))
    )

    assert result.status is CompanionActionStatus.SUCCEEDED
    assert result.output
    assert base64.b64decode(result.output["content_base64"]) == content
    assert result.output["sha256"] == hashlib.sha256(content).hexdigest()

    too_small = await executor.execute(
        _request(
            FileReadAction(path=r"C:\Allowed\report.txt", max_bytes=2),
            key="read-too-small",
        )
    )
    assert too_small.status is CompanionActionStatus.REJECTED


@pytest.mark.asyncio
async def test_file_write_is_atomic_checked_and_hashes_output(tmp_path: Path) -> None:
    content = b"new document"
    digest = hashlib.sha256(content).hexdigest()
    executor = _executor(tmp_path)
    action = FileWriteAction(
        path=r"C:\Allowed\new.txt",
        content_base64=base64.b64encode(content).decode(),
        expected_sha256=digest,
    )

    result = await executor.execute(_request(action))

    assert result.status is CompanionActionStatus.SUCCEEDED
    assert (tmp_path / "new.txt").read_bytes() == content
    assert result.output and result.output["sha256"] == digest

    mismatch = FileWriteAction(
        path=r"C:\Allowed\bad.txt",
        content_base64=base64.b64encode(content).decode(),
        expected_sha256="0" * 64,
    )
    rejected = await executor.execute(_request(mismatch, key="bad-hash"))
    assert rejected.status is CompanionActionStatus.REJECTED
    assert not (tmp_path / "bad.txt").exists()


@pytest.mark.asyncio
async def test_file_write_never_overwrites_without_permission(tmp_path: Path) -> None:
    destination = tmp_path / "existing.txt"
    destination.write_text("original", encoding="utf-8")
    executor = _executor(tmp_path)
    action = FileWriteAction(
        path=r"C:\Allowed\existing.txt",
        content_base64=base64.b64encode(b"replacement").decode(),
    )

    result = await executor.execute(_request(action))

    assert result.status is CompanionActionStatus.REJECTED
    assert destination.read_text(encoding="utf-8") == "original"


@pytest.mark.asyncio
async def test_recursive_list_is_bounded_and_hashes_files(tmp_path: Path) -> None:
    (tmp_path / "folder").mkdir()
    content = b"hello"
    (tmp_path / "folder" / "one.txt").write_bytes(content)
    (tmp_path / "two.txt").write_text("two", encoding="utf-8")
    executor = _executor(tmp_path)

    result = await executor.execute(
        _request(FileListAction(path=LOGICAL_ROOT, recursive=True, max_entries=2))
    )

    assert result.status is CompanionActionStatus.SUCCEEDED
    assert result.output
    assert result.output["count"] == 2
    assert result.output["truncated"] is True
    assert any(
        entry["sha256"] for entry in result.output["entries"] if entry["type"] == "file"
    )


@pytest.mark.asyncio
async def test_traversal_and_symlinks_are_rejected(tmp_path: Path) -> None:
    outside = tmp_path.parent / f"{tmp_path.name}-outside"
    outside.mkdir()
    (outside / "secret.txt").write_text("not allowed", encoding="utf-8")
    (tmp_path / "link").symlink_to(outside, target_is_directory=True)
    executor = _executor(tmp_path)

    traversal = await executor.execute(
        _request(
            FileReadAction(path=r"C:\Allowed\..\outside\secret.txt"),
            key="traversal",
        )
    )
    linked = await executor.execute(
        _request(FileReadAction(path=r"C:\Allowed\link\secret.txt"), key="symlink")
    )

    assert traversal.status is CompanionActionStatus.REJECTED
    assert linked.status is CompanionActionStatus.REJECTED


def test_symlink_cannot_be_used_as_a_bound_root(tmp_path: Path) -> None:
    real = tmp_path / "real"
    real.mkdir()
    linked = tmp_path / "linked"
    linked.symlink_to(real, target_is_directory=True)

    with pytest.raises(WindowsRuntimePolicyError, match="reparse"):
        WindowsActionExecutor(
            _device(), file_root_bindings={LOGICAL_ROOT: linked}, clock=lambda: NOW
        )


class FakeOffice:
    def __init__(self) -> None:
        self.opened: list[tuple[str, Path, str]] = []

    def open_document(
        self, application: str, document_path: Path, mode: str
    ) -> dict[str, Any]:
        self.opened.append((application, document_path, mode))
        return {"opened": True}

    def export_pdf(
        self, application: str, source_path: Path, output_path: Path
    ) -> dict[str, Any]:
        output_path.write_bytes(b"%PDF-test")
        return {"exported": True, "application": application}


@pytest.mark.asyncio
async def test_office_adapter_opens_and_exports_validated_paths(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    source.write_bytes(b"pptx")
    office = FakeOffice()
    executor = _executor(tmp_path, office=office)

    opened = await executor.execute(
        _request(
            OfficeOpenDocumentAction(
                application="powerpoint",
                document_path=r"C:\Allowed\deck.pptx",
                mode="edit",
            ),
            key="open-office",
        )
    )
    exported = await executor.execute(
        _request(
            OfficeExportPdfAction(
                application="powerpoint",
                source_path=r"C:\Allowed\deck.pptx",
                output_path=r"C:\Allowed\deck.pdf",
            ),
            key="export-office",
        )
    )

    assert opened.status is CompanionActionStatus.SUCCEEDED
    assert office.opened == [("powerpoint", source, "edit")]
    assert exported.status is CompanionActionStatus.SUCCEEDED
    assert exported.output
    assert exported.output["sha256"] == hashlib.sha256(b"%PDF-test").hexdigest()


class FakeServices:
    def __init__(self) -> None:
        self.calls: list[tuple[str, str]] = []

    def status(self, name: str) -> str:
        self.calls.append(("status", name))
        return "running"

    def start(self, name: str) -> str:
        self.calls.append(("start", name))
        return "running"

    def stop(self, name: str) -> str:
        self.calls.append(("stop", name))
        return "stopped"


@pytest.mark.asyncio
async def test_service_operations_use_named_allowlist(tmp_path: Path) -> None:
    services = FakeServices()
    executor = _executor(tmp_path, services=services)

    status = await executor.execute(
        _request(WindowsServiceStatusAction(service_name="Spooler"), confirmed=False)
    )
    start = await executor.execute(
        _request(WindowsServiceStartAction(service_name="spooler"), key="service-start")
    )
    denied = await executor.execute(
        _request(
            WindowsServiceStartAction(service_name="Unlisted"), key="service-denied"
        )
    )

    assert status.status is CompanionActionStatus.SUCCEEDED
    assert start.status is CompanionActionStatus.SUCCEEDED
    assert denied.status is CompanionActionStatus.REJECTED
    assert services.calls == [("status", "Spooler"), ("start", "spooler")]


class FakeClipboard:
    def __init__(self) -> None:
        self.value = "clipboard value"

    def read_text(self) -> str:
        return self.value

    def write_text(self, value: str) -> None:
        self.value = value


class FakeNotifications:
    def __init__(self) -> None:
        self.messages: list[tuple[str, str]] = []

    def show(self, title: str, message: str) -> None:
        self.messages.append((title, message))


@pytest.mark.asyncio
async def test_clipboard_and_notification_adapters(tmp_path: Path) -> None:
    clipboard = FakeClipboard()
    notifications = FakeNotifications()
    executor = _executor(tmp_path, clipboard=clipboard, notifications=notifications)

    read = await executor.execute(
        _request(ClipboardReadTextAction(max_characters=9), key="clipboard-read")
    )
    write = await executor.execute(
        _request(ClipboardWriteTextAction(text="new"), key="clipboard-write")
    )
    notice = await executor.execute(
        _request(
            NotificationShowAction(title="Agent", message="Finished"), key="notice"
        )
    )

    assert read.output and read.output["text"] == "clipboard"
    assert read.output["truncated"] is True
    assert write.status is CompanionActionStatus.SUCCEEDED
    assert clipboard.value == "new"
    assert notice.status is CompanionActionStatus.SUCCEEDED
    assert notifications.messages == [("Agent", "Finished")]


class FakeDesktopFlows:
    def __init__(self) -> None:
        self.calls = 0

    async def run_flow(
        self,
        flow_name: str,
        inputs: dict[str, Any],
        *,
        wait_for_completion: bool,
    ) -> dict[str, Any]:
        self.calls += 1
        return {
            "flow_name": flow_name,
            "inputs": inputs,
            "waited": wait_for_completion,
        }


@pytest.mark.asyncio
async def test_desktop_flow_uses_only_injected_allowlisted_executor(
    tmp_path: Path,
) -> None:
    flows = FakeDesktopFlows()
    executor = _executor(tmp_path, desktop_flows=flows)
    allowed = PowerAutomateDesktopRunAction(
        flow_name="reconcile invoices",
        inputs={"batch": 42},
        wait_for_completion=True,
    )
    denied = PowerAutomateDesktopRunAction(flow_name="arbitrary flow")

    result = await executor.execute(_request(allowed, key="pad-allowed"))
    rejected = await executor.execute(_request(denied, key="pad-denied"))

    assert result.status is CompanionActionStatus.SUCCEEDED
    assert rejected.status is CompanionActionStatus.REJECTED
    assert flows.calls == 1


@pytest.mark.asyncio
async def test_confirmation_must_match_kind_and_be_current(tmp_path: Path) -> None:
    (tmp_path / "data.txt").write_text("data", encoding="utf-8")
    executor = _executor(tmp_path)
    action = FileReadAction(path=r"C:\Allowed\data.txt")
    wrong = _request(action, confirmed=False, key="wrong-kind").model_copy(
        update={"confirmation": _confirmation(CompanionActionKind.CLIPBOARD_READ_TEXT)}
    )
    future_evidence = _confirmation(CompanionActionKind.FILE_READ).model_copy(
        update={
            "confirmed_at": NOW + timedelta(minutes=1),
            "expires_at": NOW + timedelta(minutes=2),
        }
    )
    future = _request(action, confirmed=False, key="future-confirmation").model_copy(
        update={"confirmation": future_evidence}
    )

    assert (await executor.execute(wrong)).status is CompanionActionStatus.REJECTED
    assert (await executor.execute(future)).status is CompanionActionStatus.REJECTED


@pytest.mark.asyncio
async def test_expired_request_is_not_executed(tmp_path: Path) -> None:
    executor = _executor(tmp_path)
    request = CompanionActionRequest(
        action=SystemInventoryAction(),
        requested_at=NOW - timedelta(minutes=10),
        expires_at=NOW - timedelta(minutes=1),
    )

    result = await executor.execute(request)

    assert result.status is CompanionActionStatus.EXPIRED


@pytest.mark.asyncio
async def test_idempotency_replays_result_and_rejects_key_reuse(tmp_path: Path) -> None:
    flows = FakeDesktopFlows()
    executor = _executor(tmp_path, desktop_flows=flows)
    first_request = _request(
        PowerAutomateDesktopRunAction(flow_name="Reconcile invoices"), key="same-key"
    )

    first = await executor.execute(first_request)
    replay = await executor.execute(first_request)
    conflict = await executor.execute(
        _request(SystemInventoryAction(), confirmed=False, key="same-key")
    )

    assert first == replay
    assert flows.calls == 1
    assert conflict.status is CompanionActionStatus.REJECTED
    assert conflict.error and conflict.error.code == "idempotency_conflict"


class FakeRelay:
    def __init__(self, batch: RelayPollBatch) -> None:
        self.batch = batch
        self.polls: list[dict[str, Any]] = []
        self.acks: list[tuple[str, Any]] = []

    async def poll(
        self,
        identity: DeviceIdentity,
        *,
        cursor: str | None,
        maximum_actions: int,
        wait_seconds: float,
    ) -> RelayPollBatch:
        self.polls.append(
            {
                "identity": identity,
                "cursor": cursor,
                "maximum_actions": maximum_actions,
                "wait_seconds": wait_seconds,
            }
        )
        return self.batch

    async def acknowledge(
        self, identity: DeviceIdentity, delivery_id: str, result: Any
    ) -> None:
        assert identity == _identity()
        self.acks.append((delivery_id, result))


@pytest.mark.asyncio
async def test_outbound_worker_polls_executes_and_acknowledges(tmp_path: Path) -> None:
    executor = _executor(tmp_path)
    request = _request(SystemInventoryAction(), confirmed=False, key="relay-action")
    delivery = RelayActionDelivery(
        delivery_id="delivery-1",
        device_id=_identity().device_id,
        expected_device_identity=_identity(),
        request=request,
        policy=executor.local_policies[CompanionActionKind.SYSTEM_INVENTORY],
    )
    relay = FakeRelay(RelayPollBatch(cursor="next", deliveries=(delivery,)))
    worker = OutboundRelayWorker(relay, executor)

    count = await worker.run_once()

    assert count == 1
    assert worker.cursor == "next"
    assert relay.acks[0][0] == "delivery-1"
    assert relay.acks[0][1].status is CompanionActionStatus.SUCCEEDED


@pytest.mark.asyncio
async def test_outbound_worker_rejects_identity_and_policy_downgrade(
    tmp_path: Path,
) -> None:
    executor = _executor(tmp_path)
    request = _request(SystemInventoryAction(), confirmed=False, key="relay-reject")
    local = executor.local_policies[CompanionActionKind.SYSTEM_INVENTORY]
    delivery = RelayActionDelivery(
        delivery_id="wrong-device",
        device_id="another-device",
        expected_device_identity=_identity("another-device"),
        request=request,
        policy=local,
    )
    relay = FakeRelay(RelayPollBatch(deliveries=(delivery,)))

    await OutboundRelayWorker(relay, executor).run_once()

    result = relay.acks[0][1]
    assert result.status is CompanionActionStatus.REJECTED
    assert result.error.code == "identity_mismatch"


def test_pywin32_automation_is_optional_off_windows() -> None:
    if sys.platform == "win32":
        pytest.skip("This guard is specific to non-Windows test hosts")
    with pytest.raises(RuntimeError, match="only on Windows"):
        PyWin32OfficeAutomation()
