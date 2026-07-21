"""Unit tests for the supported Power Platform integration module."""

from __future__ import annotations

import asyncio
import json
from collections.abc import Mapping
from typing import Any
from uuid import UUID, uuid4

import pytest
from pydantic import ValidationError

from microsoft_agent.power_platform import (
    DesktopFlowConnectionType,
    DesktopFlowRunMode,
    DesktopFlowSchemaKind,
    FlowState,
    HttpResponse,
    NamedDesktopFlow,
    NamedFlowTrigger,
    PowerPlatformClient,
    PowerPlatformClientError,
    PowerPlatformErrorCode,
    PowerPlatformSettings,
)

FLOW_ID = UUID("11111111-1111-4111-8111-111111111111")
FLOW_ID_UNIQUE = UUID("22222222-2222-4222-8222-222222222222")
DESKTOP_FLOW_ID = UUID("33333333-3333-4333-8333-333333333333")
FLOW_SESSION_ID = UUID("44444444-4444-4444-8444-444444444444")


def _flow(**overrides: Any) -> dict[str, Any]:
    value: dict[str, Any] = {
        "workflowid": str(FLOW_ID),
        "workflowidunique": str(FLOW_ID_UNIQUE),
        "name": "Document approval",
        "description": "Approve generated documents",
        "category": 5,
        "statecode": 1,
        "type": 1,
        "ismanaged": False,
        "createdon": "2026-01-01T00:00:00Z",
        "modifiedon": "2026-02-01T00:00:00Z",
    }
    value.update(overrides)
    return value


def _response(
    payload: Any = None,
    *,
    status: int = 200,
    headers: Mapping[str, str] | None = None,
) -> HttpResponse:
    body = b"" if payload is None else json.dumps(payload).encode()
    return HttpResponse(status_code=status, headers=dict(headers or {}), body=body)


class FakeTransport:
    def __init__(self, *responses: HttpResponse | Exception) -> None:
        self.responses = list(responses)
        self.calls: list[dict[str, Any]] = []

    async def request(
        self,
        method: str,
        url: str,
        *,
        headers: Mapping[str, str],
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        timeout: float,
    ) -> HttpResponse:
        self.calls.append(
            {
                "method": method,
                "url": url,
                "headers": dict(headers),
                "params": params,
                "body": body,
                "timeout": timeout,
            }
        )
        result = self.responses.pop(0)
        if isinstance(result, Exception):
            raise result
        return result


class FakeTokenProvider:
    def __init__(self, token: str = "token", *, delay: float = 0) -> None:
        self.token = token
        self.delay = delay
        self.audiences: list[str] = []

    async def get_token(self, audience: str) -> str:
        self.audiences.append(audience)
        if self.delay:
            await asyncio.sleep(self.delay)
        return self.token


def _settings(**overrides: Any) -> PowerPlatformSettings:
    values: dict[str, Any] = {
        "dataverse_environment_url": "https://contoso.crm.dynamics.com",
    }
    values.update(overrides)
    return PowerPlatformSettings(**values)


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("dataverse_environment_url", "http://contoso.crm.dynamics.com"),
        ("dataverse_environment_url", "https://api.flow.microsoft.com"),
        (
            "named_flows",
            {"bad": {"trigger_url": "https://api.flow.microsoft.com/providers/foo"}},
        ),
    ],
)
def test_settings_reject_insecure_or_unsupported_endpoints(
    field: str, value: Any
) -> None:
    values: dict[str, Any] = {
        "dataverse_environment_url": "https://contoso.crm.dynamics.com",
        field: value,
    }
    with pytest.raises(ValidationError):
        PowerPlatformSettings(**values)


@pytest.mark.asyncio
async def test_list_solution_flows_uses_documented_dataverse_query() -> None:
    transport = FakeTransport(_response({"value": [_flow()]}))
    tokens = FakeTokenProvider()
    client = PowerPlatformClient(_settings(), tokens, transport)

    result = await client.list_solution_flows(active_only=True, top=25)

    assert result.flows[0].workflow_id == FLOW_ID
    assert result.flows[0].state is FlowState.ACTIVATED
    assert result.pages_fetched == 1
    call = transport.calls[0]
    assert call["url"] == ("https://contoso.crm.dynamics.com/api/data/v9.2/workflows")
    assert call["params"]["$filter"] == (
        "category eq 5 and type eq 1 and statecode eq 1"
    )
    assert call["params"]["$top"] == 25
    assert call["headers"]["Authorization"] == "Bearer token"
    assert tokens.audiences == ["https://contoso.crm.dynamics.com"]


@pytest.mark.asyncio
async def test_list_solution_flows_follows_same_origin_next_link() -> None:
    next_link = (
        "https://contoso.crm.dynamics.com/api/data/v9.2/workflows?$skiptoken=abc"
    )
    second = _flow(
        workflowid=str(UUID("33333333-3333-4333-8333-333333333333")),
        name="Second flow",
    )
    transport = FakeTransport(
        _response({"value": [_flow()], "@odata.nextLink": next_link}),
        _response({"value": [second]}),
    )
    client = PowerPlatformClient(_settings(), FakeTokenProvider(), transport)

    result = await client.list_solution_flows(fetch_all=True)

    assert len(result.flows) == 2
    assert result.pages_fetched == 2
    assert transport.calls[1]["url"] == next_link
    assert transport.calls[1]["params"] is None


@pytest.mark.asyncio
async def test_list_solution_flows_rejects_cross_origin_next_link() -> None:
    transport = FakeTransport(
        _response(
            {
                "value": [_flow()],
                "@odata.nextLink": "https://attacker.invalid/steal-token",
            }
        )
    )
    client = PowerPlatformClient(_settings(), FakeTokenProvider(), transport)

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.list_solution_flows(fetch_all=True)

    assert exc_info.value.error.code is PowerPlatformErrorCode.INVALID_RESPONSE
    assert len(transport.calls) == 1


@pytest.mark.asyncio
async def test_get_solution_flow_rejects_non_modern_workflow() -> None:
    transport = FakeTransport(_response(_flow(category=0)))
    client = PowerPlatformClient(_settings(), FakeTokenProvider(), transport)

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.get_solution_flow(FLOW_ID)

    assert exc_info.value.error.code is PowerPlatformErrorCode.INVALID_RESPONSE


@pytest.mark.asyncio
async def test_lifecycle_changes_are_deny_by_default() -> None:
    transport = FakeTransport()
    tokens = FakeTokenProvider()
    client = PowerPlatformClient(_settings(), tokens, transport)

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.activate_solution_flow(FLOW_ID)

    assert exc_info.value.error.code is PowerPlatformErrorCode.POLICY
    assert not transport.calls
    assert not tokens.audiences


@pytest.mark.asyncio
async def test_activate_solution_flow_patches_state_with_concurrency_headers() -> None:
    transport = FakeTransport(_response(status=204))
    client = PowerPlatformClient(
        _settings(allow_lifecycle_changes=True), FakeTokenProvider(), transport
    )

    result = await client.activate_solution_flow(
        FLOW_ID, idempotency_key="activate-approval-v1", etag='W/"42"'
    )

    assert result.state is FlowState.ACTIVATED
    call = transport.calls[0]
    assert call["method"] == "PATCH"
    assert call["url"].endswith(f"/workflows({FLOW_ID})")
    assert json.loads(call["body"]) == {"statecode": 1}
    assert call["headers"]["If-Match"] == 'W/"42"'
    assert call["headers"]["Idempotency-Key"] == "activate-approval-v1"


@pytest.mark.asyncio
async def test_named_flow_trigger_is_oauth_authenticated_and_idempotent() -> None:
    trigger = NamedFlowTrigger(
        trigger_url=(
            "https://prod-00.westus.logic.azure.com/workflows/abc/"
            "triggers/manual/paths/invoke"
        ),
        workflow_id=FLOW_ID,
    )
    transport = FakeTransport(
        _response(
            {"run_id": "run-123"},
            status=202,
            headers={
                "Content-Type": "application/json",
                "Location": "https://status.example/run-123",
            },
        )
    )
    tokens = FakeTokenProvider()
    client = PowerPlatformClient(
        _settings(named_flows={"document_approval": trigger}), tokens, transport
    )

    result = await client.trigger_solution_flow(
        "document_approval",
        {"document_id": "doc-1"},
        workflow_id=FLOW_ID,
        idempotency_key="doc-1-approval",
    )

    assert result.accepted is True
    assert result.output == {"run_id": "run-123"}
    assert result.location == "https://status.example/run-123"
    assert tokens.audiences == ["https://service.flow.microsoft.com/"]
    call = transport.calls[0]
    assert call["headers"]["Authorization"] == "Bearer token"
    assert call["headers"]["Idempotency-Key"] == "doc-1-approval"
    assert json.loads(call["body"]) == {"document_id": "doc-1"}


@pytest.mark.asyncio
async def test_unknown_named_flow_is_rejected_before_token_or_transport() -> None:
    transport = FakeTransport()
    tokens = FakeTokenProvider()
    client = PowerPlatformClient(_settings(), tokens, transport)

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.trigger_solution_flow("raw-url-is-not-accepted", {})

    assert exc_info.value.error.code is PowerPlatformErrorCode.POLICY
    assert not transport.calls
    assert not tokens.audiences


@pytest.mark.asyncio
async def test_named_flow_workflow_binding_is_enforced() -> None:
    trigger = NamedFlowTrigger(
        trigger_url="https://logic.example.com/triggers/manual/invoke",
        workflow_id=FLOW_ID,
    )
    client = PowerPlatformClient(
        _settings(named_flows={"approval": trigger}),
        FakeTokenProvider(),
        FakeTransport(),
    )

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.trigger_solution_flow("approval", {}, workflow_id=uuid4())

    assert exc_info.value.error.code is PowerPlatformErrorCode.POLICY


@pytest.mark.asyncio
async def test_upstream_error_is_safely_shaped() -> None:
    transport = FakeTransport(
        _response(
            {"error": {"code": "0x800723", "message": "Service protection"}},
            status=429,
            headers={"Retry-After": "17", "x-ms-request-id": "request-1"},
        )
    )
    client = PowerPlatformClient(_settings(), FakeTokenProvider(), transport)

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.list_solution_flows()

    error = exc_info.value.error
    assert error.code is PowerPlatformErrorCode.RATE_LIMITED
    assert error.upstream_code == "0x800723"
    assert error.message == "Service protection"
    assert error.retry_after_seconds == 17
    assert error.correlation_id == "request-1"


@pytest.mark.asyncio
async def test_token_timeout_is_normalized() -> None:
    client = PowerPlatformClient(
        _settings(timeout_seconds=0.01),
        FakeTokenProvider(delay=0.1),
        FakeTransport(),
    )

    with pytest.raises(PowerPlatformClientError) as exc_info:
        await client.list_solution_flows()

    assert exc_info.value.error.code is PowerPlatformErrorCode.TIMEOUT


@pytest.mark.asyncio
async def test_list_and_schema_desktop_flows_use_documented_dataverse_paths() -> None:
    desktop = NamedDesktopFlow(
        workflow_id=DESKTOP_FLOW_ID,
        connection_name="shared-uiflow-connection",
    )
    transport = FakeTransport(
        _response(
            {
                "value": [
                    {
                        "workflowid": str(DESKTOP_FLOW_ID),
                        "name": "Prepare report",
                        "category": 6,
                    }
                ]
            }
        ),
        _response({"schema": {"type": "object", "properties": {}}}),
    )
    client = PowerPlatformClient(
        _settings(named_desktop_flows={"prepare_report": desktop}),
        FakeTokenProvider(),
        transport,
    )

    listed = await client.list_desktop_flows(include_unpublished=True)
    schema = await client.get_desktop_flow_schema(
        "prepare_report", DesktopFlowSchemaKind.INPUTS
    )

    assert listed.flows[0].workflow_id == DESKTOP_FLOW_ID
    assert transport.calls[0]["params"]["$filter"] == "category eq 6"
    assert transport.calls[0]["headers"]["MSCRM.IncludeUnpublished"] == "true"
    assert transport.calls[1]["url"].endswith(
        f"/workflows({DESKTOP_FLOW_ID})/inputs/$value"
    )
    assert schema["schema"]["type"] == "object"


@pytest.mark.asyncio
async def test_desktop_flow_run_status_outputs_and_cancel_are_functional() -> None:
    desktop = NamedDesktopFlow(
        workflow_id=DESKTOP_FLOW_ID,
        connection_name="desktop-connection-reference",
        connection_type=DesktopFlowConnectionType.CONNECTION_REFERENCE,
        allowed_run_modes=frozenset({DesktopFlowRunMode.UNATTENDED}),
        timeout_seconds=7200,
    )
    transport = FakeTransport(
        _response({"flowsessionId": str(FLOW_SESSION_ID)}),
        _response(
            {
                "statuscode": 8,
                "statecode": 0,
                "startedon": "2026-07-17T12:00:00Z",
                "completedon": "2026-07-17T12:02:00Z",
            }
        ),
        _response({"ReportPath": r"C:\Reports\result.pdf"}),
        _response(status=204),
    )
    client = PowerPlatformClient(
        _settings(
            named_desktop_flows={"prepare_report": desktop},
            allow_desktop_flow_runs=True,
            allow_desktop_flow_cancellations=True,
        ),
        FakeTokenProvider(),
        transport,
    )

    queued = await client.run_desktop_flow(
        "prepare_report",
        {"ReportId": "report-1"},
        run_mode=DesktopFlowRunMode.UNATTENDED,
        timeout_seconds=3600,
    )
    status = await client.get_desktop_flow_run_status(queued.flow_session_id)
    outputs = await client.get_desktop_flow_outputs(queued.flow_session_id)
    cancelled = await client.cancel_desktop_flow_run(queued.flow_session_id)

    run_call = transport.calls[0]
    assert run_call["url"].endswith(
        f"/workflows({DESKTOP_FLOW_ID})/Microsoft.Dynamics.CRM.RunDesktopFlow"
    )
    body = json.loads(run_call["body"])
    assert body["connectionName"] == "desktop-connection-reference"
    assert body["connectionType"] == 2
    assert body["runMode"] == "unattended"
    assert json.loads(body["inputs"]) == {"ReportId": "report-1"}
    assert queued.flow_session_id == FLOW_SESSION_ID
    assert status.status_code == 8
    assert outputs == {"ReportPath": r"C:\Reports\result.pdf"}
    assert cancelled.flow_session_id == FLOW_SESSION_ID
    assert transport.calls[3]["url"].endswith(
        f"/flowsessions({FLOW_SESSION_ID})/Microsoft.Dynamics.CRM.CancelDesktopFlowRun"
    )


@pytest.mark.asyncio
async def test_desktop_flow_mutations_are_fail_closed() -> None:
    desktop = NamedDesktopFlow(
        workflow_id=DESKTOP_FLOW_ID,
        connection_name="desktop-connection",
    )
    transport = FakeTransport()
    client = PowerPlatformClient(
        _settings(named_desktop_flows={"prepare_report": desktop}),
        FakeTokenProvider(),
        transport,
    )

    with pytest.raises(PowerPlatformClientError) as run_error:
        await client.run_desktop_flow("prepare_report", {})
    with pytest.raises(PowerPlatformClientError) as cancel_error:
        await client.cancel_desktop_flow_run(FLOW_SESSION_ID)

    assert run_error.value.error.code is PowerPlatformErrorCode.POLICY
    assert cancel_error.value.error.code is PowerPlatformErrorCode.POLICY
    assert not transport.calls
