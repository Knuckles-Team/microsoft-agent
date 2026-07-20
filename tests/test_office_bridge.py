"""Focused tests for the typed Office.js relay and its trust boundaries."""

from __future__ import annotations

import json
from datetime import UTC, datetime, timedelta
from typing import Any

import pytest
from fastmcp import FastMCP
from starlette.requests import Request

from microsoft_agent.office_bridge import (
    OfficeBridgeError,
    OfficeBridgeStore,
    OfficeCapabilityReport,
    OfficeCommandKind,
    OfficeCommandState,
    OfficeCommandSucceeded,
    OfficeHost,
    OfficePairSessionRequest,
    OfficeRequirementSupport,
    PowerPointAddSlideCommand,
    PowerPointAddSlideResult,
    WordReadSelectionCommand,
    WordSelectionResult,
    register_office_bridge,
)
from microsoft_agent.settings import clear_settings_cache


class MutableClock:
    """Small deterministic aware clock for expiry tests."""

    def __init__(self) -> None:
        self.value = datetime(2026, 7, 17, 12, tzinfo=UTC)

    def __call__(self) -> datetime:
        return self.value

    def advance(self, delta: timedelta) -> None:
        self.value += delta


def capabilities(host: OfficeHost = OfficeHost.WORD) -> OfficeCapabilityReport:
    requirement_set = "WordApi" if host is OfficeHost.WORD else "PowerPointApi"
    return OfficeCapabilityReport(
        host=host,
        platform="PC",
        office_version="16.0",
        requirements=(
            OfficeRequirementSupport(
                requirement_set=requirement_set,
                version="1.1",
                supported=True,
            ),
        ),
    )


async def paired_session(store: OfficeBridgeStore, host: OfficeHost = OfficeHost.WORD):
    pairing = await store.create_pairing(host, "Quarterly planning")
    session = await store.pair_session(
        OfficePairSessionRequest(
            pairing_token=pairing.pairing_token,
            capabilities=capabilities(host),
        )
    )
    return pairing, session


@pytest.mark.asyncio
async def test_pairing_is_one_time_host_bound_and_not_stored_in_plaintext() -> None:
    store = OfficeBridgeStore()
    pairing = await store.create_pairing(OfficeHost.WORD, "Document A")

    assert pairing.pairing_token not in repr(store.__dict__)
    with pytest.raises(OfficeBridgeError, match="another Office host"):
        await store.pair_session(
            OfficePairSessionRequest(
                pairing_token=pairing.pairing_token,
                capabilities=capabilities(OfficeHost.POWERPOINT),
            )
        )

    session = await store.pair_session(
        OfficePairSessionRequest(
            pairing_token=pairing.pairing_token,
            capabilities=capabilities(),
        )
    )
    assert session.session_token not in repr(store.__dict__)
    with pytest.raises(OfficeBridgeError, match="invalid or expired"):
        await store.pair_session(
            OfficePairSessionRequest(
                pairing_token=pairing.pairing_token,
                capabilities=capabilities(),
            )
        )


@pytest.mark.asyncio
async def test_command_round_trip_is_typed_and_bound_to_one_session() -> None:
    store = OfficeBridgeStore()
    _, session = await paired_session(store)
    receipt = await store.enqueue(session.session_id, WordReadSelectionCommand())

    command = await store.poll(session.session_token, wait_seconds=0)
    assert command is not None
    assert command.command_id == receipt.command_id
    assert command.payload.kind is OfficeCommandKind.WORD_READ_SELECTION

    outcome = OfficeCommandSucceeded(
        command_id=command.command_id,
        status="succeeded",
        kind=OfficeCommandKind.WORD_READ_SELECTION,
        result=WordSelectionResult(
            kind=OfficeCommandKind.WORD_READ_SELECTION,
            text="selected words",
        ),
    )
    completed = await store.complete(session.session_token, outcome)

    assert completed.state is OfficeCommandState.SUCCEEDED
    assert completed.outcome == outcome
    assert (await store.wait_for_result(command.command_id, 0)) == completed


@pytest.mark.asyncio
async def test_bridge_rejects_wrong_host_and_wrong_session_results() -> None:
    store = OfficeBridgeStore()
    _, word_session = await paired_session(store, OfficeHost.WORD)
    _, powerpoint_session = await paired_session(store, OfficeHost.POWERPOINT)

    with pytest.raises(OfficeBridgeError, match="does not match"):
        await store.enqueue(word_session.session_id, PowerPointAddSlideCommand())

    receipt = await store.enqueue(
        powerpoint_session.session_id, PowerPointAddSlideCommand()
    )
    command = await store.poll(powerpoint_session.session_token, wait_seconds=0)
    assert command is not None
    outcome = OfficeCommandSucceeded(
        command_id=receipt.command_id,
        status="succeeded",
        kind=OfficeCommandKind.POWERPOINT_ADD_SLIDE,
        result=PowerPointAddSlideResult(
            kind=OfficeCommandKind.POWERPOINT_ADD_SLIDE,
            slide={"number": 2, "id": "slide-2"},
        ),
    )
    with pytest.raises(OfficeBridgeError, match="not found"):
        await store.complete(word_session.session_token, outcome)


@pytest.mark.asyncio
async def test_pairing_session_and_command_expiry_are_enforced() -> None:
    clock = MutableClock()
    store = OfficeBridgeStore(clock=clock)
    pairing = await store.create_pairing(OfficeHost.WORD, "Expiring")
    clock.advance(timedelta(minutes=6))
    with pytest.raises(OfficeBridgeError, match="invalid or expired"):
        await store.pair_session(
            OfficePairSessionRequest(
                pairing_token=pairing.pairing_token,
                capabilities=capabilities(),
            )
        )

    fresh = await store.create_pairing(OfficeHost.WORD, "Command expiry")
    session = await store.pair_session(
        OfficePairSessionRequest(
            pairing_token=fresh.pairing_token,
            capabilities=capabilities(),
        )
    )
    receipt = await store.enqueue(session.session_id, WordReadSelectionCommand())
    clock.advance(timedelta(minutes=3))
    expired = await store.get_command(receipt.command_id)
    assert expired.state is OfficeCommandState.EXPIRED

    clock.advance(timedelta(minutes=13))
    assert await store.list_sessions() == ()
    with pytest.raises(OfficeBridgeError, match="not found"):
        await store.get_command(receipt.command_id)
    with pytest.raises(OfficeBridgeError, match="invalid"):
        await store.poll(session.session_token, wait_seconds=0)


@pytest.mark.asyncio
async def test_capacity_is_bounded_per_session() -> None:
    store = OfficeBridgeStore(max_commands_per_session=1)
    _, session = await paired_session(store)
    await store.enqueue(session.session_id, WordReadSelectionCommand())

    with pytest.raises(OfficeBridgeError, match="capacity"):
        await store.enqueue(session.session_id, WordReadSelectionCommand())


def make_request(
    method: str,
    path: str,
    origin: str,
    *,
    body: dict[str, Any] | None = None,
    token: str | None = None,
) -> Request:
    encoded = json.dumps(body or {}).encode("utf-8")
    headers = [
        (b"origin", origin.encode("ascii")),
        (b"content-type", b"application/json"),
        (b"content-length", str(len(encoded)).encode("ascii")),
    ]
    if token is not None:
        headers.append((b"authorization", f"Bearer {token}".encode("ascii")))
    delivered = False

    async def receive() -> dict[str, Any]:
        nonlocal delivered
        if delivered:
            return {"type": "http.disconnect"}
        delivered = True
        return {"type": "http.request", "body": encoded, "more_body": False}

    return Request(
        {
            "type": "http",
            "http_version": "1.1",
            "method": method,
            "scheme": "https",
            "path": path,
            "raw_path": path.encode("ascii"),
            "query_string": b"",
            "headers": headers,
            "client": ("127.0.0.1", 12345),
            "server": ("localhost", 8000),
        },
        receive,
    )


def route_endpoint(mcp: FastMCP, path: str):
    return next(
        route.endpoint for route in mcp._additional_http_routes if route.path == path
    )


@pytest.mark.asyncio
async def test_http_routes_require_exact_cors_origin_and_redeem_pairing(
    monkeypatch,
) -> None:
    monkeypatch.setenv("MICROSOFT_OFFICE_ADDIN_ORIGINS", "https://office.example.test")
    clear_settings_cache()
    try:
        store = OfficeBridgeStore()
        mcp = FastMCP("office-bridge-test")
        register_office_bridge(mcp, store)
        endpoint = route_endpoint(mcp, "/office-bridge/session")
        denied = await endpoint(
            make_request(
                "OPTIONS",
                "/office-bridge/session",
                "https://office.example.test.evil.test",
            )
        )
        assert denied.status_code == 403

        pairing = await store.create_pairing(OfficeHost.WORD, "Browser session")
        response = await endpoint(
            make_request(
                "POST",
                "/office-bridge/session",
                "https://office.example.test",
                body={
                    "pairing_token": pairing.pairing_token,
                    "capabilities": capabilities().model_dump(mode="json"),
                },
            )
        )
    finally:
        clear_settings_cache()

    assert response.status_code == 201
    assert response.headers["access-control-allow-origin"] == (
        "https://office.example.test"
    )
    parsed = json.loads(response.body)
    assert parsed["host"] == "Word"
    assert len(parsed["session_token"]) >= 32


@pytest.mark.asyncio
async def test_http_poll_and_result_routes_complete_a_typed_command(
    monkeypatch,
) -> None:
    monkeypatch.setenv("MICROSOFT_OFFICE_ADDIN_ORIGINS", "https://office.example.test")
    clear_settings_cache()
    try:
        store = OfficeBridgeStore()
        _, session = await paired_session(store)
        receipt = await store.enqueue(session.session_id, WordReadSelectionCommand())
        mcp = FastMCP("office-bridge-test")
        register_office_bridge(mcp, store)

        poll_response = await route_endpoint(mcp, "/office-bridge/commands/poll")(
            make_request(
                "POST",
                "/office-bridge/commands/poll",
                "https://office.example.test",
                body={"wait_seconds": 0},
                token=session.session_token,
            )
        )
        command = json.loads(poll_response.body)["command"]
        assert command["command_id"] == str(receipt.command_id)

        result_response = await route_endpoint(mcp, "/office-bridge/commands/result")(
            make_request(
                "POST",
                "/office-bridge/commands/result",
                "https://office.example.test",
                body={
                    "command_id": command["command_id"],
                    "status": "succeeded",
                    "kind": "word.read_selection",
                    "result": {
                        "kind": "word.read_selection",
                        "text": "typed response",
                    },
                },
                token=session.session_token,
            )
        )
    finally:
        clear_settings_cache()

    assert result_response.status_code == 200
    assert json.loads(result_response.body)["state"] == "succeeded"
    assert (await store.get_command(receipt.command_id)).state is (
        OfficeCommandState.SUCCEEDED
    )


@pytest.mark.asyncio
async def test_registration_exposes_individual_policy_classifiable_tools() -> None:
    mcp = FastMCP("office-bridge-test")
    register_office_bridge(mcp, OfficeBridgeStore())

    names = {tool.name for tool in await mcp.list_tools(run_middleware=False)}
    assert {
        "create_office_pairing",
        "list_office_sessions",
        "get_word_selection_from_office",
        "write_word_selection_in_office",
        "delete_powerpoint_slide_in_office",
    } <= names
