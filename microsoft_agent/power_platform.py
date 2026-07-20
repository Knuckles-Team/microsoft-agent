"""Supported Power Platform and Power Automate integration primitives.

Solution-aware cloud flows are managed through the documented Dataverse Web
API ``workflows`` entity set.  Flow execution is deliberately separate:
callers configure named, OAuth-protected "When an HTTP request is received"
endpoints and invoke them by name.  This module never uses the unsupported
``api.flow.microsoft.com`` management API and never accepts a trigger URL from
an invocation request.

The client is independent of any particular OAuth library. Production uses the
shared Agent Utilities TLS and DNS-pinned HTTP boundary; tests may inject the
minimal transport protocol directly.
"""

from __future__ import annotations

import asyncio
import json
import re
from collections.abc import Mapping
from datetime import UTC, datetime
from enum import IntEnum, StrEnum
from typing import Any, Never, Protocol, runtime_checkable
from urllib.parse import urlparse
from uuid import UUID, uuid4

import httpx
from agent_utilities.core.http_client import create_http_client
from agent_utilities.core.transport_security import (
    ResolvedTLSProfile,
    resolve_configured_tls_profile,
)
from agent_utilities.security.cli_secrets import validate_runtime_secret_reference
from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    HttpUrl,
    field_validator,
    model_validator,
)

PUBLIC_FLOW_SERVICE_AUDIENCE = "https://service.flow.microsoft.com/"
_UNSUPPORTED_FLOW_API_HOST = "api.flow.microsoft.com"
_JSON_HEADERS = {"Accept": "application/json", "Content-Type": "application/json"}
_ODATA_HEADERS = {
    **_JSON_HEADERS,
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": 'odata.include-annotations="*"',
}
_FLOW_SELECT = ",".join(
    (
        "category",
        "createdon",
        "description",
        "ismanaged",
        "modifiedon",
        "name",
        "statecode",
        "type",
        "workflowid",
        "workflowidunique",
        "_createdby_value",
        "_modifiedby_value",
        "_ownerid_value",
    )
)


class HttpResponse(BaseModel):
    """Transport-neutral HTTP response."""

    model_config = ConfigDict(frozen=True)

    status_code: int = Field(ge=100, le=599)
    headers: dict[str, str] = Field(default_factory=dict)
    body: bytes = b""

    def json_body(self) -> Any:
        """Decode the response body as JSON."""

        if not self.body:
            return None
        return json.loads(self.body.decode("utf-8"))


@runtime_checkable
class AsyncHttpTransport(Protocol):
    """Minimal injectable async HTTP transport."""

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
        """Send one HTTP request without automatically following redirects."""


@runtime_checkable
class AudienceTokenProvider(Protocol):
    """Acquire an OAuth access token for a resource audience."""

    async def get_token(self, audience: str) -> str:
        """Return a bearer token whose ``aud`` claim matches ``audience``."""


class HttpxAsyncHttpTransport:
    """Bounded async facade over the shared verified sync HTTP boundary."""

    def __init__(
        self,
        *,
        service: str,
        tls_profile: str | None = None,
        tls_profile_ref: str | None = None,
        allowed_private_hosts: tuple[str, ...] = (),
        max_response_bytes: int = 16 * 1024 * 1024,
        client: httpx.Client | None = None,
    ) -> None:
        if not 1_024 <= max_response_bytes <= 64 * 1024 * 1024:
            raise ValueError("HTTP response bound is invalid")
        self._max_response_bytes = max_response_bytes
        self._tls: ResolvedTLSProfile | None = None
        self._owns_client = client is None
        self._closed = False
        if client is not None:
            self._client = client
            return
        self._tls = resolve_configured_tls_profile(
            service,
            profile_name=tls_profile,
            profile_ref=tls_profile_ref,
        )
        if self._tls.proxy_url:
            self._tls.cleanup()
            self._tls = None
            raise ValueError("Pinned provider transport does not support a proxy")
        try:
            self._client = create_http_client(
                timeout=httpx.Timeout(30.0),
                verify=self._tls.ssl_context,
                follow_redirects=False,
                trust_env=False,
                pin_egress=True,
                allowed_private_hosts=allowed_private_hosts,
                limits=httpx.Limits(
                    max_connections=32,
                    max_keepalive_connections=8,
                ),
            )
        except Exception:
            self._tls.cleanup()
            self._tls = None
            raise

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
        """Send one request without following redirects."""

        try:
            response = await asyncio.to_thread(
                self._client.request,
                method,
                url,
                headers=dict(headers),
                params=params,
                content=body,
                timeout=timeout,
                follow_redirects=False,
            )
        except httpx.TimeoutException:
            raise TimeoutError("Provider request timed out") from None
        except httpx.TransportError:
            raise OSError("Provider transport failed") from None
        if len(response.content) > self._max_response_bytes:
            raise ValueError("Provider response exceeds the configured bound")
        return HttpResponse(
            status_code=response.status_code,
            headers=dict(response.headers),
            body=response.content,
        )

    def close(self) -> None:
        """Close owned HTTP state and remove materialized trust files."""
        if self._closed:
            return
        self._closed = True
        try:
            if self._owns_client:
                self._client.close()
        finally:
            if self._tls is not None:
                self._tls.cleanup()
                self._tls = None


class FlowState(IntEnum):
    """Documented state values for modern flows in Dataverse."""

    DRAFT = 0
    ACTIVATED = 1
    SUSPENDED = 2


class FlowType(IntEnum):
    """Dataverse workflow record types."""

    DEFINITION = 1
    ACTIVATION = 2
    TEMPLATE = 3


class DesktopFlowRunMode(StrEnum):
    """Documented execution modes for a Power Automate desktop flow."""

    ATTENDED = "attended"
    UNATTENDED = "unattended"


class DesktopFlowPriority(StrEnum):
    """Documented desktop-flow queue priorities."""

    NORMAL = "normal"
    HIGH = "high"


class DesktopFlowSchemaKind(StrEnum):
    """Published schema documents exposed by a desktop-flow workflow."""

    INPUTS = "inputs"
    OUTPUTS = "outputs"


class DesktopFlowConnectionType(IntEnum):
    """Connection identifier types accepted by ``RunDesktopFlow``."""

    CONNECTION = 1
    CONNECTION_REFERENCE = 2


class NamedDesktopFlow(BaseModel):
    """Allowlisted desktop flow and its fixed machine connection binding."""

    model_config = ConfigDict(frozen=True)

    workflow_id: UUID
    connection_name: str = Field(min_length=1, max_length=256)
    connection_type: DesktopFlowConnectionType = DesktopFlowConnectionType.CONNECTION
    allowed_run_modes: frozenset[DesktopFlowRunMode] = frozenset(
        {DesktopFlowRunMode.ATTENDED}
    )
    timeout_seconds: int = Field(default=10_800, ge=1, le=86_400)
    enabled: bool = True

    @field_validator("connection_name")
    @classmethod
    def validate_connection_name(cls, value: str) -> str:
        if value != value.strip() or any(ord(char) < 32 for char in value):
            raise ValueError("desktop-flow connection names must be trimmed text")
        return value

    @field_validator("allowed_run_modes")
    @classmethod
    def validate_run_modes(
        cls, value: frozenset[DesktopFlowRunMode]
    ) -> frozenset[DesktopFlowRunMode]:
        if not value:
            raise ValueError("desktop flows need at least one allowed run mode")
        return value


class NamedFlowTrigger(BaseModel):
    """An allowlisted OAuth HTTP trigger, selected by its mapping key."""

    model_config = ConfigDict(frozen=True)

    trigger_url: HttpUrl
    audience: str = PUBLIC_FLOW_SERVICE_AUDIENCE
    workflow_id: UUID | None = None
    enabled: bool = True
    timeout_seconds: float | None = Field(default=None, gt=0, le=300)
    expected_status_codes: frozenset[int] = Field(
        default_factory=lambda: frozenset({200, 201, 202, 204})
    )

    @field_validator("trigger_url")
    @classmethod
    def validate_trigger_url(cls, value: HttpUrl) -> HttpUrl:
        """Require TLS and reject Microsoft's unsupported management host."""

        if value.scheme != "https":
            raise ValueError("flow trigger URLs must use HTTPS")
        if value.username or value.password:
            raise ValueError("flow trigger URLs cannot contain user credentials")
        if (value.host or "").lower() == _UNSUPPORTED_FLOW_API_HOST:
            raise ValueError("api.flow.microsoft.com is not a supported API")
        return value

    @field_validator("audience")
    @classmethod
    def validate_audience(cls, value: str) -> str:
        value = value.strip()
        parsed = urlparse(value)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("OAuth audience must be an absolute HTTPS resource")
        return value

    @field_validator("expected_status_codes")
    @classmethod
    def validate_expected_status_codes(cls, value: frozenset[int]) -> frozenset[int]:
        if not value or any(code < 200 or code > 299 for code in value):
            raise ValueError("expected status codes must be non-empty and successful")
        return value


class PowerPlatformSettings(BaseModel):
    """Non-secret configuration for one Dataverse environment."""

    model_config = ConfigDict(frozen=True)

    dataverse_environment_url: HttpUrl
    dataverse_audience: str | None = None
    api_version: str = Field(default="v9.2", pattern=r"^v\d+\.\d+$")
    timeout_seconds: float = Field(default=30.0, gt=0, le=300)
    max_pages: int = Field(default=20, ge=1, le=100)
    allow_lifecycle_changes: bool = False
    allow_desktop_flow_runs: bool = False
    allow_desktop_flow_cancellations: bool = False
    tls_profile: str | None = None
    tls_profile_ref: str | None = None
    named_flows: dict[str, NamedFlowTrigger] = Field(default_factory=dict)
    named_desktop_flows: dict[str, NamedDesktopFlow] = Field(default_factory=dict)

    @field_validator("dataverse_environment_url")
    @classmethod
    def validate_environment_url(cls, value: HttpUrl) -> HttpUrl:
        if value.scheme != "https":
            raise ValueError("Dataverse environment URL must use HTTPS")
        if value.username or value.password:
            raise ValueError("Dataverse environment URL cannot contain credentials")
        if (value.host or "").lower() == _UNSUPPORTED_FLOW_API_HOST:
            raise ValueError("api.flow.microsoft.com is not a Dataverse environment")
        if value.path not in {"", "/"} or value.query or value.fragment:
            raise ValueError("use the Dataverse organization root URL without a path")
        return value

    @field_validator("dataverse_audience")
    @classmethod
    def validate_dataverse_audience(cls, value: str | None) -> str | None:
        if value is None:
            return None
        value = value.strip()
        parsed = urlparse(value)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("Dataverse audience must be an absolute HTTPS resource")
        return value

    @field_validator("named_flows", "named_desktop_flows")
    @classmethod
    def validate_flow_names(cls, value: dict[str, Any]) -> dict[str, Any]:
        for name in value:
            if not name or name != name.strip() or len(name) > 128:
                raise ValueError("named flow keys must be trimmed and 1-128 characters")
        return value

    @field_validator("tls_profile")
    @classmethod
    def validate_tls_profile(cls, value: str | None) -> str | None:
        if value is None:
            return None
        selected = value.strip()
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9_.-]{0,127}", selected) is None:
            raise ValueError("Power Platform TLS profile name is invalid")
        return selected

    @field_validator("tls_profile_ref")
    @classmethod
    def validate_tls_profile_ref(cls, value: str | None) -> str | None:
        if value is None:
            return None
        return validate_runtime_secret_reference(value)

    @model_validator(mode="after")
    def validate_tls_selector(self) -> PowerPlatformSettings:
        if self.tls_profile and self.tls_profile_ref:
            raise ValueError("Power Platform TLS selectors are ambiguous")
        return self

    @property
    def api_base_url(self) -> str:
        """Return the documented Dataverse Web API root."""

        return (
            f"{str(self.dataverse_environment_url).rstrip('/')}"
            f"/api/data/{self.api_version}"
        )

    @property
    def token_audience(self) -> str:
        """Return the audience used when acquiring a Dataverse token."""

        return self.dataverse_audience or str(self.dataverse_environment_url).rstrip(
            "/"
        )


class FlowRecord(BaseModel):
    """Supported subset of a solution-aware Dataverse workflow row."""

    model_config = ConfigDict(populate_by_name=True, extra="allow")

    workflow_id: UUID = Field(alias="workflowid")
    workflow_id_unique: UUID | None = Field(default=None, alias="workflowidunique")
    name: str
    description: str | None = None
    category: int
    state: FlowState = Field(alias="statecode")
    flow_type: FlowType = Field(alias="type")
    is_managed: bool | None = Field(default=None, alias="ismanaged")
    created_on: datetime | None = Field(default=None, alias="createdon")
    modified_on: datetime | None = Field(default=None, alias="modifiedon")
    owner_id: UUID | None = Field(default=None, alias="_ownerid_value")
    etag: str | None = Field(default=None, alias="@odata.etag")

    @model_validator(mode="after")
    def require_modern_flow(self) -> FlowRecord:
        if self.category != 5:
            raise ValueError("workflow row is not a modern cloud flow (category 5)")
        return self


class FlowListResult(BaseModel):
    """A bounded result from the Dataverse workflow collection."""

    model_config = ConfigDict(frozen=True)

    flows: list[FlowRecord]
    next_link: str | None = None
    pages_fetched: int = Field(ge=1)


class FlowLifecycleResult(BaseModel):
    """Confirmation that Dataverse accepted a lifecycle update."""

    model_config = ConfigDict(frozen=True)

    workflow_id: UUID
    state: FlowState
    changed_at: datetime
    request_id: UUID
    idempotency_key: str


class FlowInvocationResult(BaseModel):
    """Normalized result from an allowlisted flow HTTP trigger."""

    model_config = ConfigDict(frozen=True)

    flow_name: str
    workflow_id: UUID | None = None
    status_code: int
    accepted: bool
    request_id: UUID
    idempotency_key: str
    location: str | None = None
    output: Any = None


class DesktopFlowRecord(BaseModel):
    """Published or draft Power Automate desktop-flow definition."""

    model_config = ConfigDict(populate_by_name=True, extra="allow")

    workflow_id: UUID = Field(alias="workflowid")
    name: str
    category: int = 6
    etag: str | None = Field(default=None, alias="@odata.etag")

    @model_validator(mode="after")
    def require_desktop_flow(self) -> DesktopFlowRecord:
        if self.category != 6:
            raise ValueError("workflow row is not a desktop flow (category 6)")
        return self


class DesktopFlowListResult(BaseModel):
    """Bounded list of desktop flows visible in one Dataverse environment."""

    model_config = ConfigDict(frozen=True)

    flows: list[DesktopFlowRecord]
    next_link: str | None = None
    pages_fetched: int = Field(ge=1)


class DesktopFlowRunResult(BaseModel):
    """Dataverse acknowledgement for a newly queued desktop-flow run."""

    model_config = ConfigDict(frozen=True)

    flow_name: str
    workflow_id: UUID
    flow_session_id: UUID
    run_mode: DesktopFlowRunMode
    accepted_at: datetime


class DesktopFlowRunStatus(BaseModel):
    """Documented status fields for a Dataverse desktop-flow session."""

    model_config = ConfigDict(populate_by_name=True, extra="allow")

    flow_session_id: UUID
    status_code: int = Field(alias="statuscode")
    state_code: int = Field(alias="statecode")
    started_on: datetime | None = Field(default=None, alias="startedon")
    completed_on: datetime | None = Field(default=None, alias="completedon")


class DesktopFlowCancellationResult(BaseModel):
    """Confirmation that Dataverse accepted a desktop-flow cancellation."""

    model_config = ConfigDict(frozen=True)

    flow_session_id: UUID
    cancelled_at: datetime


class PowerPlatformErrorCode(StrEnum):
    """Stable error categories exposed to MCP tools and agents."""

    POLICY = "policy_denied"
    AUTHENTICATION = "authentication_failed"
    FORBIDDEN = "forbidden"
    NOT_FOUND = "not_found"
    CONFLICT = "conflict"
    RATE_LIMITED = "rate_limited"
    TIMEOUT = "timeout"
    TRANSPORT = "transport_error"
    INVALID_RESPONSE = "invalid_response"
    UPSTREAM = "upstream_error"


class PowerPlatformError(BaseModel):
    """Safe, normalized error information."""

    model_config = ConfigDict(frozen=True)

    code: PowerPlatformErrorCode
    message: str
    status_code: int | None = None
    upstream_code: str | None = None
    retry_after_seconds: int | None = None
    correlation_id: str | None = None


class PowerPlatformClientError(RuntimeError):
    """Raised for policy, transport, and upstream failures."""

    def __init__(self, error: PowerPlatformError):
        self.error = error
        super().__init__(error.message)


class PowerPlatformClient:
    """Manage solution-aware flows and invoke allowlisted OAuth triggers."""

    def __init__(
        self,
        settings: PowerPlatformSettings,
        token_provider: AudienceTokenProvider,
        transport: AsyncHttpTransport,
    ) -> None:
        self.settings = settings
        self._token_provider = token_provider
        self._transport = transport
        # Snapshot the validated allowlist so later mutation of the source dict
        # cannot add an endpoint to a live client.
        self._named_flows = dict(settings.named_flows)
        self._named_desktop_flows = dict(settings.named_desktop_flows)

    async def list_solution_flows(
        self,
        *,
        active_only: bool = False,
        top: int = 100,
        fetch_all: bool = False,
    ) -> FlowListResult:
        """List modern cloud flows stored in the Dataverse workflow table."""

        if top < 1 or top > 5000:
            raise ValueError("top must be between 1 and 5000")
        filters = ["category eq 5", "type eq 1"]
        if active_only:
            filters.append("statecode eq 1")
        params: Mapping[str, Any] | None = {
            "$filter": " and ".join(filters),
            "$select": _FLOW_SELECT,
            "$orderby": "modifiedon desc",
            "$top": top,
        }
        url = f"{self.settings.api_base_url}/workflows"
        flows: list[FlowRecord] = []
        pages_fetched = 0
        next_link: str | None = None

        while True:
            response = await self._dataverse_request("GET", url, params=params)
            payload = self._require_json_object(response)
            raw_flows = payload.get("value")
            if not isinstance(raw_flows, list):
                self._raise_invalid_response("Dataverse response has no value array")
            try:
                flows.extend(FlowRecord.model_validate(item) for item in raw_flows)
            except (TypeError, ValueError):
                self._raise_invalid_response("Invalid Dataverse flow row")
            pages_fetched += 1
            raw_next = payload.get("@odata.nextLink")
            next_link = raw_next if isinstance(raw_next, str) and raw_next else None
            if not fetch_all or not next_link:
                break
            if pages_fetched >= self.settings.max_pages:
                break
            self._validate_dataverse_next_link(next_link)
            url = next_link
            params = None

        return FlowListResult(
            flows=flows,
            next_link=next_link,
            pages_fetched=pages_fetched,
        )

    async def get_solution_flow(self, workflow_id: UUID) -> FlowRecord:
        """Retrieve one modern cloud flow from Dataverse."""

        url = f"{self.settings.api_base_url}/workflows({workflow_id})"
        response = await self._dataverse_request(
            "GET", url, params={"$select": _FLOW_SELECT}
        )
        try:
            return FlowRecord.model_validate(self._require_json_object(response))
        except (TypeError, ValueError):
            self._raise_invalid_response("Invalid Dataverse flow row")

    async def list_desktop_flows(
        self,
        *,
        top: int = 100,
        include_unpublished: bool = False,
        fetch_all: bool = False,
    ) -> DesktopFlowListResult:
        """List Power Automate desktop flows through the Dataverse Web API."""

        if top < 1 or top > 5000:
            raise ValueError("top must be between 1 and 5000")
        params: Mapping[str, Any] | None = {
            "$filter": "category eq 6",
            "$select": "name,workflowid,category",
            "$orderby": "name",
            "$top": top,
        }
        url = f"{self.settings.api_base_url}/workflows"
        flows: list[DesktopFlowRecord] = []
        pages_fetched = 0
        next_link: str | None = None
        extra_headers = (
            {"MSCRM.IncludeUnpublished": "true"} if include_unpublished else None
        )

        while True:
            response = await self._dataverse_request(
                "GET", url, params=params, extra_headers=extra_headers
            )
            payload = self._require_json_object(response)
            rows = payload.get("value")
            if not isinstance(rows, list):
                self._raise_invalid_response("Dataverse response has no value array")
            try:
                flows.extend(DesktopFlowRecord.model_validate(item) for item in rows)
            except (TypeError, ValueError):
                self._raise_invalid_response("Invalid Dataverse desktop-flow row")
            pages_fetched += 1
            raw_next = payload.get("@odata.nextLink")
            next_link = raw_next if isinstance(raw_next, str) and raw_next else None
            if (
                not fetch_all
                or not next_link
                or pages_fetched >= self.settings.max_pages
            ):
                break
            self._validate_dataverse_next_link(next_link)
            url = next_link
            params = None

        return DesktopFlowListResult(
            flows=flows,
            next_link=next_link,
            pages_fetched=pages_fetched,
        )

    async def get_desktop_flow_schema(
        self, flow_name: str, kind: DesktopFlowSchemaKind
    ) -> dict[str, Any]:
        """Get the documented input or output schema for an allowlisted desktop flow."""

        configured = self._desktop_flow(flow_name)
        url = (
            f"{self.settings.api_base_url}/workflows({configured.workflow_id})/"
            f"{kind.value}/$value"
        )
        response = await self._dataverse_request("GET", url)
        return self._require_json_object(response)

    async def run_desktop_flow(
        self,
        flow_name: str,
        inputs: Mapping[str, Any] | None = None,
        *,
        run_mode: DesktopFlowRunMode = DesktopFlowRunMode.ATTENDED,
        priority: DesktopFlowPriority = DesktopFlowPriority.NORMAL,
        timeout_seconds: int | None = None,
    ) -> DesktopFlowRunResult:
        """Queue an allowlisted desktop flow with a fixed connection binding."""

        if not self.settings.allow_desktop_flow_runs:
            self._raise_policy("Desktop-flow execution is disabled")
        configured = self._desktop_flow(flow_name)
        if run_mode not in configured.allowed_run_modes:
            self._raise_policy(
                f"Run mode {run_mode.value!r} is not allowed for flow {flow_name!r}"
            )
        timeout = (
            configured.timeout_seconds if timeout_seconds is None else timeout_seconds
        )
        if timeout < 1 or timeout > configured.timeout_seconds:
            self._raise_policy(
                "Desktop-flow timeout cannot exceed the configured per-flow maximum"
            )
        serialized_inputs = self._json_bytes(dict(inputs or {})).decode("utf-8")
        if len(serialized_inputs.encode("utf-8")) > 2 * 1024 * 1024:
            raise ValueError("desktop-flow inputs exceed the Dataverse 2 MiB limit")
        url = (
            f"{self.settings.api_base_url}/workflows({configured.workflow_id})/"
            "Microsoft.Dynamics.CRM.RunDesktopFlow"
        )
        response = await self._dataverse_request(
            "POST",
            url,
            body=self._json_bytes(
                {
                    "connectionName": configured.connection_name,
                    "connectionType": int(configured.connection_type),
                    "inputs": serialized_inputs,
                    "runMode": run_mode.value,
                    "runPriority": priority.value,
                    "timeout": timeout,
                }
            ),
            expected_statuses=frozenset({200}),
        )
        payload = self._require_json_object(response)
        try:
            session_id = UUID(str(payload["flowsessionId"]))
        except (KeyError, TypeError, ValueError):
            self._raise_invalid_response(
                "RunDesktopFlow response has no valid flowsessionId"
            )
        return DesktopFlowRunResult(
            flow_name=flow_name,
            workflow_id=configured.workflow_id,
            flow_session_id=session_id,
            run_mode=run_mode,
            accepted_at=datetime.now(UTC),
        )

    async def get_desktop_flow_run_status(
        self, flow_session_id: UUID
    ) -> DesktopFlowRunStatus:
        """Get the documented state and timestamps for a desktop-flow session."""

        url = f"{self.settings.api_base_url}/flowsessions({flow_session_id})"
        response = await self._dataverse_request(
            "GET",
            url,
            params={"$select": "statuscode,statecode,startedon,completedon"},
        )
        payload = self._require_json_object(response)
        try:
            return DesktopFlowRunStatus.model_validate(
                {**payload, "flow_session_id": flow_session_id}
            )
        except (TypeError, ValueError):
            self._raise_invalid_response("Invalid desktop-flow session row")

    async def get_desktop_flow_outputs(self, flow_session_id: UUID) -> dict[str, Any]:
        """Read outputs for a completed desktop-flow session."""

        url = f"{self.settings.api_base_url}/flowsessions({flow_session_id})/outputs/$value"
        response = await self._dataverse_request("GET", url)
        return self._require_json_object(response)

    async def cancel_desktop_flow_run(
        self, flow_session_id: UUID
    ) -> DesktopFlowCancellationResult:
        """Cancel a queued or running desktop flow when explicitly enabled."""

        if not self.settings.allow_desktop_flow_cancellations:
            self._raise_policy("Desktop-flow cancellation is disabled")
        url = (
            f"{self.settings.api_base_url}/flowsessions({flow_session_id})/"
            "Microsoft.Dynamics.CRM.CancelDesktopFlowRun"
        )
        await self._dataverse_request("POST", url, expected_statuses=frozenset({204}))
        return DesktopFlowCancellationResult(
            flow_session_id=flow_session_id,
            cancelled_at=datetime.now(UTC),
        )

    async def activate_solution_flow(
        self,
        workflow_id: UUID,
        *,
        idempotency_key: str | None = None,
        etag: str = "*",
    ) -> FlowLifecycleResult:
        """Turn on a modern cloud flow using a Dataverse row update."""

        return await self.set_solution_flow_state(
            workflow_id,
            FlowState.ACTIVATED,
            idempotency_key=idempotency_key,
            etag=etag,
        )

    async def deactivate_solution_flow(
        self,
        workflow_id: UUID,
        *,
        idempotency_key: str | None = None,
        etag: str = "*",
    ) -> FlowLifecycleResult:
        """Turn off a modern cloud flow using a Dataverse row update."""

        return await self.set_solution_flow_state(
            workflow_id,
            FlowState.DRAFT,
            idempotency_key=idempotency_key,
            etag=etag,
        )

    async def set_solution_flow_state(
        self,
        workflow_id: UUID,
        state: FlowState,
        *,
        idempotency_key: str | None = None,
        etag: str = "*",
    ) -> FlowLifecycleResult:
        """Set a supported flow lifecycle state through Dataverse."""

        if not self.settings.allow_lifecycle_changes:
            self._raise_policy(
                "Flow lifecycle changes are disabled by PowerPlatformSettings"
            )
        if state not in {FlowState.DRAFT, FlowState.ACTIVATED}:
            self._raise_policy("Only activate and deactivate operations are supported")
        key = self._idempotency_key(idempotency_key)
        request_id = uuid4()
        url = f"{self.settings.api_base_url}/workflows({workflow_id})"
        await self._dataverse_request(
            "PATCH",
            url,
            body=self._json_bytes({"statecode": int(state)}),
            extra_headers={
                "If-Match": etag,
                "Idempotency-Key": key,
                "x-ms-client-request-id": str(request_id),
            },
            expected_statuses=frozenset({204}),
        )
        return FlowLifecycleResult(
            workflow_id=workflow_id,
            state=state,
            changed_at=datetime.now(UTC),
            request_id=request_id,
            idempotency_key=key,
        )

    async def trigger_solution_flow(
        self,
        flow_name: str,
        payload: Mapping[str, Any] | None = None,
        *,
        workflow_id: UUID | None = None,
        idempotency_key: str | None = None,
        timeout_seconds: float | None = None,
    ) -> FlowInvocationResult:
        """Invoke a named OAuth HTTP trigger bound to a solution flow.

        The endpoint is resolved exclusively from ``settings.named_flows``.
        Supplying a URL dynamically is intentionally impossible.
        """

        configured = self._named_flows.get(flow_name)
        if configured is None:
            self._raise_policy(f"Flow {flow_name!r} is not allowlisted")
        if not configured.enabled:
            self._raise_policy(f"Flow {flow_name!r} is disabled")
        if (
            workflow_id is not None
            and configured.workflow_id is not None
            and configured.workflow_id != workflow_id
        ):
            self._raise_policy(
                "Configured workflow ID does not match the requested flow"
            )

        key = self._idempotency_key(idempotency_key)
        request_id = uuid4()
        timeout = timeout_seconds or configured.timeout_seconds
        response = await self._request(
            "POST",
            str(configured.trigger_url),
            audience=configured.audience,
            headers={
                **_JSON_HEADERS,
                "Idempotency-Key": key,
                "x-ms-client-request-id": str(request_id),
            },
            body=self._json_bytes(dict(payload or {})),
            timeout=timeout,
            expected_statuses=configured.expected_status_codes,
        )
        output: Any = None
        if response.body:
            content_type = self._header(response.headers, "content-type") or ""
            if "json" in content_type.lower():
                try:
                    output = response.json_body()
                except (UnicodeDecodeError, json.JSONDecodeError):
                    self._raise_invalid_response("Flow trigger returned invalid JSON")
            else:
                output = response.body.decode("utf-8", errors="replace")
        return FlowInvocationResult(
            flow_name=flow_name,
            workflow_id=configured.workflow_id,
            status_code=response.status_code,
            accepted=True,
            request_id=request_id,
            idempotency_key=key,
            location=self._header(response.headers, "location"),
            output=output,
        )

    def _desktop_flow(self, flow_name: str) -> NamedDesktopFlow:
        configured = self._named_desktop_flows.get(flow_name)
        if configured is None:
            self._raise_policy(f"Desktop flow {flow_name!r} is not allowlisted")
        if not configured.enabled:
            self._raise_policy(f"Desktop flow {flow_name!r} is disabled")
        return configured

    async def _dataverse_request(
        self,
        method: str,
        url: str,
        *,
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        extra_headers: Mapping[str, str] | None = None,
        expected_statuses: frozenset[int] = frozenset({200}),
    ) -> HttpResponse:
        headers = {**_ODATA_HEADERS, **dict(extra_headers or {})}
        return await self._request(
            method,
            url,
            audience=self.settings.token_audience,
            headers=headers,
            params=params,
            body=body,
            expected_statuses=expected_statuses,
        )

    async def _request(
        self,
        method: str,
        url: str,
        *,
        audience: str,
        headers: Mapping[str, str],
        params: Mapping[str, Any] | None = None,
        body: bytes | None = None,
        timeout: float | None = None,
        expected_statuses: frozenset[int],
    ) -> HttpResponse:
        request_timeout = timeout or self.settings.timeout_seconds
        try:
            token = await asyncio.wait_for(
                self._token_provider.get_token(audience), timeout=request_timeout
            )
        except TimeoutError as exc:
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.TIMEOUT,
                    message=(
                        "Power Platform token acquisition timed out after "
                        f"{request_timeout:g}s"
                    ),
                )
            ) from exc
        except Exception as exc:
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.AUTHENTICATION,
                    message=(
                        f"Power Platform token acquisition failed: {type(exc).__name__}"
                    ),
                )
            ) from exc
        if not token or not token.strip():
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.AUTHENTICATION,
                    message=f"Token provider returned no token for {audience}",
                )
            )

        try:
            response = await asyncio.wait_for(
                self._transport.request(
                    method,
                    url,
                    headers={**dict(headers), "Authorization": f"Bearer {token}"},
                    params=params,
                    body=body,
                    timeout=request_timeout,
                ),
                timeout=request_timeout,
            )
        except TimeoutError as exc:
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.TIMEOUT,
                    message=f"Power Platform request timed out after {request_timeout:g}s",
                )
            ) from exc
        except Exception as exc:
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.TRANSPORT,
                    message=f"Power Platform transport failed: {type(exc).__name__}",
                )
            ) from exc

        if response.status_code not in expected_statuses:
            raise PowerPlatformClientError(self._error_from_response(response))
        return response

    def _validate_dataverse_next_link(self, next_link: str) -> None:
        parsed = urlparse(next_link)
        api = urlparse(self.settings.api_base_url)
        if (
            parsed.scheme != "https"
            or parsed.hostname != api.hostname
            or parsed.port != api.port
            or not parsed.path.startswith(f"{api.path}/")
        ):
            self._raise_invalid_response(
                "Dataverse nextLink changed origin or API root"
            )

    @staticmethod
    def _json_bytes(value: Any) -> bytes:
        try:
            return json.dumps(
                value,
                ensure_ascii=False,
                separators=(",", ":"),
                sort_keys=True,
            ).encode("utf-8")
        except (TypeError, ValueError) as exc:
            raise ValueError("request payload must be JSON serializable") from exc

    @staticmethod
    def _require_json_object(response: HttpResponse) -> dict[str, Any]:
        try:
            payload = response.json_body()
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.INVALID_RESPONSE,
                    message="Power Platform returned invalid JSON",
                    status_code=response.status_code,
                )
            ) from exc
        if not isinstance(payload, dict):
            raise PowerPlatformClientError(
                PowerPlatformError(
                    code=PowerPlatformErrorCode.INVALID_RESPONSE,
                    message="Power Platform response must be a JSON object",
                    status_code=response.status_code,
                )
            )
        return payload

    @classmethod
    def _error_from_response(cls, response: HttpResponse) -> PowerPlatformError:
        status = response.status_code
        code = PowerPlatformErrorCode.UPSTREAM
        if status == 401:
            code = PowerPlatformErrorCode.AUTHENTICATION
        elif status == 403:
            code = PowerPlatformErrorCode.FORBIDDEN
        elif status == 404:
            code = PowerPlatformErrorCode.NOT_FOUND
        elif status in {409, 412}:
            code = PowerPlatformErrorCode.CONFLICT
        elif status == 429:
            code = PowerPlatformErrorCode.RATE_LIMITED

        message = f"Power Platform request failed with HTTP {status}"
        upstream_code: str | None = None
        if response.body:
            try:
                payload = response.json_body()
                error = payload.get("error") if isinstance(payload, dict) else None
                if isinstance(error, dict):
                    raw_code = error.get("code")
                    raw_message = error.get("message")
                    if isinstance(raw_code, str):
                        upstream_code = raw_code[:256]
                    if isinstance(raw_message, str) and raw_message.strip():
                        message = raw_message.strip()[:1000]
            except (UnicodeDecodeError, json.JSONDecodeError):
                pass

        retry_after: int | None = None
        raw_retry = cls._header(response.headers, "retry-after")
        if raw_retry and raw_retry.isdigit():
            retry_after = int(raw_retry)
        correlation = (
            cls._header(response.headers, "x-ms-request-id")
            or cls._header(response.headers, "request-id")
            or cls._header(response.headers, "x-ms-correlation-request-id")
        )
        return PowerPlatformError(
            code=code,
            message=message,
            status_code=status,
            upstream_code=upstream_code,
            retry_after_seconds=retry_after,
            correlation_id=correlation,
        )

    @staticmethod
    def _header(headers: Mapping[str, str], name: str) -> str | None:
        wanted = name.casefold()
        for key, value in headers.items():
            if key.casefold() == wanted:
                return value
        return None

    @staticmethod
    def _idempotency_key(value: str | None) -> str:
        key = value or str(uuid4())
        if not 1 <= len(key) <= 128 or any(
            char.isspace() or ord(char) < 33 or ord(char) > 126 for char in key
        ):
            raise ValueError(
                "idempotency key must contain 1-128 visible ASCII characters"
            )
        return key

    @staticmethod
    def _raise_policy(message: str) -> Never:
        raise PowerPlatformClientError(
            PowerPlatformError(code=PowerPlatformErrorCode.POLICY, message=message)
        )

    @staticmethod
    def _raise_invalid_response(message: str) -> Never:
        raise PowerPlatformClientError(
            PowerPlatformError(
                code=PowerPlatformErrorCode.INVALID_RESPONSE,
                message=message,
            )
        )


__all__ = [
    "AsyncHttpTransport",
    "AudienceTokenProvider",
    "DesktopFlowCancellationResult",
    "DesktopFlowConnectionType",
    "DesktopFlowListResult",
    "DesktopFlowPriority",
    "DesktopFlowRecord",
    "DesktopFlowRunMode",
    "DesktopFlowRunResult",
    "DesktopFlowRunStatus",
    "DesktopFlowSchemaKind",
    "FlowInvocationResult",
    "FlowLifecycleResult",
    "FlowListResult",
    "FlowRecord",
    "FlowState",
    "FlowType",
    "HttpResponse",
    "HttpxAsyncHttpTransport",
    "NamedFlowTrigger",
    "NamedDesktopFlow",
    "PUBLIC_FLOW_SERVICE_AUDIENCE",
    "PowerPlatformClient",
    "PowerPlatformClientError",
    "PowerPlatformError",
    "PowerPlatformErrorCode",
    "PowerPlatformSettings",
]
