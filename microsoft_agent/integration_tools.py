"""FastMCP registration for Office, Power Platform, and Windows integrations."""

from __future__ import annotations

from functools import lru_cache
from typing import Any, TypeVar
from urllib.parse import urlparse
from uuid import UUID

from fastmcp import FastMCP
from pydantic import Field

from microsoft_agent.document_service import (
    ArtifactDelivery,
    ArtifactOptions,
    DocumentService,
    PowerPointPresentationRequest,
    WordDocumentRequest,
    get_document_capabilities,
)
from microsoft_agent.graph_file_service import (
    GRAPH_AUDIENCE,
    GraphFileService,
    GraphFileSettings,
    UploadConflictBehavior,
)
from microsoft_agent.integration_adapters import (
    IntuneGraphTokenAdapter,
    IntuneHttpClientAdapter,
)
from microsoft_agent.integration_auth import MicrosoftAudienceTokenProvider
from microsoft_agent.integration_config import get_integration_settings
from microsoft_agent.intune_service import (
    ConfirmationEvidence as IntuneConfirmationEvidence,
)
from microsoft_agent.intune_service import IntuneService
from microsoft_agent.power_platform import (
    DesktopFlowPriority,
    DesktopFlowRunMode,
    DesktopFlowSchemaKind,
    FlowState,
    HttpxAsyncHttpTransport,
    PowerPlatformClient,
)
from microsoft_agent.settings import get_settings
from microsoft_agent.windows_companion import (
    CompanionAction,
    CompanionActionKind,
    CompanionActionRequest,
    ConfirmationEvidence,
    HttpxCompanionTransport,
    WindowsCompanionClient,
)

_ManagedTransport = HttpxAsyncHttpTransport | HttpxCompanionTransport
_ManagedTransportT = TypeVar("_ManagedTransportT", bound=_ManagedTransport)
_managed_transports: list[_ManagedTransport] = []


@lru_cache(maxsize=1)
def _document_service() -> DocumentService:
    settings = get_integration_settings().documents
    if settings.artifact_root is None:
        raise ValueError(
            "Document generation requires a deployment-owned artifact root"
        )
    return DocumentService(
        artifact_root=settings.artifact_root,
        template_root=settings.template_root,
        max_artifact_bytes=settings.max_artifact_bytes,
        max_template_bytes=settings.max_template_bytes,
    )


def _configured_hosts(*urls: Any) -> tuple[str, ...]:
    """Return exact authorities derived only from validated runtime settings."""

    return tuple(
        dict.fromkeys(
            host
            for value in urls
            if (host := (urlparse(str(value)).hostname or "").casefold())
        )
    )


def _manage_transport(transport: _ManagedTransportT) -> _ManagedTransportT:
    _managed_transports.append(transport)
    return transport


def _http_transport(
    *,
    service: str,
    tls_profile: str | None = None,
    tls_profile_ref: str | None = None,
    allowed_private_hosts: tuple[str, ...] = (),
) -> HttpxAsyncHttpTransport:
    return _manage_transport(
        HttpxAsyncHttpTransport(
            service=service,
            tls_profile=tls_profile,
            tls_profile_ref=tls_profile_ref,
            allowed_private_hosts=allowed_private_hosts,
        )
    )


def _companion_transport(
    *,
    tls_profile: str | None,
    tls_profile_ref: str | None,
    allowed_private_hosts: tuple[str, ...],
) -> HttpxCompanionTransport:
    return _manage_transport(
        HttpxCompanionTransport(
            tls_profile=tls_profile,
            tls_profile_ref=tls_profile_ref,
            allowed_private_hosts=allowed_private_hosts,
        )
    )


@lru_cache(maxsize=1)
def _graph_file_service() -> GraphFileService:
    settings = get_settings()
    provider = MicrosoftAudienceTokenProvider([GRAPH_AUDIENCE])
    return GraphFileService(
        GraphFileSettings(graph_base_url=settings.graph_base_url),
        provider,
        _http_transport(
            service="microsoft_graph",
            tls_profile=settings.graph_tls_profile,
            tls_profile_ref=settings.graph_tls_profile_ref,
        ),
    )


@lru_cache(maxsize=1)
def _power_platform_client() -> PowerPlatformClient:
    settings = get_integration_settings().power_platform
    if settings is None:
        raise ValueError(
            "Power Platform is not configured. Set "
            "MICROSOFT_DATAVERSE_ENVIRONMENT_URL or provide power_platform in "
            "MICROSOFT_INTEGRATIONS_CONFIG_PATH."
        )
    audiences = {settings.token_audience}
    audiences.update(str(flow.audience) for flow in settings.named_flows.values())
    return PowerPlatformClient(
        settings,
        MicrosoftAudienceTokenProvider(audiences),
        _http_transport(
            service="microsoft_power_platform",
            tls_profile=settings.tls_profile,
            tls_profile_ref=settings.tls_profile_ref,
            allowed_private_hosts=_configured_hosts(
                settings.dataverse_environment_url,
                *(flow.trigger_url for flow in settings.named_flows.values()),
            ),
        ),
    )


def _power_platform_audiences() -> set[str]:
    settings = get_integration_settings().power_platform
    if settings is None:
        raise ValueError("Power Platform is not configured.")
    audiences = {settings.token_audience}
    audiences.update(str(flow.audience) for flow in settings.named_flows.values())
    return audiences


@lru_cache(maxsize=1)
def _windows_companion_client() -> WindowsCompanionClient:
    settings = get_integration_settings().windows_companion
    if settings is None:
        raise ValueError(
            "Windows companion is not configured. Add windows_companion to "
            "MICROSOFT_INTEGRATIONS_CONFIG_PATH."
        )
    return WindowsCompanionClient(
        settings,
        MicrosoftAudienceTokenProvider([settings.token_audience]),
        _companion_transport(
            tls_profile=settings.tls_profile,
            tls_profile_ref=settings.tls_profile_ref,
            allowed_private_hosts=_configured_hosts(settings.control_plane_url),
        ),
    )


@lru_cache(maxsize=1)
def _intune_service() -> IntuneService:
    settings = get_integration_settings().intune
    if settings is None:
        raise ValueError(
            "Intune is not configured. Add intune device/action allowlists to "
            "MICROSOFT_INTEGRATIONS_CONFIG_PATH."
        )
    provider = MicrosoftAudienceTokenProvider([GRAPH_AUDIENCE])
    graph_settings = get_settings()
    return IntuneService(
        settings,
        IntuneHttpClientAdapter(
            _http_transport(
                service="microsoft_graph",
                tls_profile=graph_settings.graph_tls_profile,
                tls_profile_ref=graph_settings.graph_tls_profile_ref,
            )
        ),
        IntuneGraphTokenAdapter(provider),
    )


def clear_integration_client_caches() -> None:
    """Close owned transports, then clear immutable integration singletons."""

    failures: list[Exception] = []
    try:
        while _managed_transports:
            transport = _managed_transports.pop()
            try:
                transport.close()
            except Exception as exc:  # pragma: no cover - defensive shutdown path
                failures.append(exc)
    finally:
        _document_service.cache_clear()
        _graph_file_service.cache_clear()
        _power_platform_client.cache_clear()
        _windows_companion_client.cache_clear()
        _intune_service.cache_clear()
    if failures:
        raise RuntimeError(
            "Microsoft integration transport cleanup failed"
        ) from failures[0]


def register_document_tools(mcp: FastMCP) -> None:
    """Register safe Word and PowerPoint generation and upload tools."""

    @mcp.tool(
        name="get_document_capabilities",
        description=(
            "Report whether local Word and PowerPoint OOXML generation backends "
            "are installed. This tool does not require Microsoft authentication."
        ),
        tags={"documents", "word", "powerpoint"},
        annotations={"readOnlyHint": True},
    )
    async def document_capabilities() -> dict[str, Any]:
        """Return document backend availability and configured trust boundaries."""

        capabilities = get_document_capabilities().model_dump(mode="json")
        settings = get_integration_settings().documents
        capabilities.update(
            {
                "artifact_root_configured": bool(settings.artifact_root),
                "template_root_configured": bool(settings.template_root),
                "max_artifact_bytes": settings.max_artifact_bytes,
            }
        )
        return capabilities

    @mcp.tool(
        name="generate_word_document",
        description=(
            "Generate a non-macro Word .docx file from validated paragraphs, "
            "tables, metadata, or a template confined to the configured root."
        ),
        tags={"documents", "word"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def create_word_document(request: WordDocumentRequest) -> dict[str, Any]:
        """Generate a Word document and return bytes, base64, or a confined file."""

        artifact = await _document_service().generate_word_document(request)
        return artifact.model_dump(mode="json")

    @mcp.tool(
        name="generate_powerpoint_presentation",
        description=(
            "Generate a non-macro PowerPoint .pptx presentation from validated "
            "slides, metadata, or a template confined to the configured root."
        ),
        tags={"documents", "powerpoint"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def create_powerpoint_presentation(
        request: PowerPointPresentationRequest,
    ) -> dict[str, Any]:
        """Generate a PowerPoint presentation and return its selected delivery."""

        artifact = await _document_service().generate_powerpoint_presentation(request)
        return artifact.model_dump(mode="json")

    @mcp.tool(
        name="generate_and_upload_word_document",
        description=(
            "Generate a Word .docx document in memory and upload it to a "
            "OneDrive or SharePoint document-library drive path."
        ),
        tags={"documents", "word", "files", "sharepoint"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def create_and_upload_word_document(
        request: WordDocumentRequest,
        drive_id: str = Field(..., description="Target OneDrive/SharePoint drive ID"),
        destination_path: str = Field(
            ..., description="Relative destination path including .docx filename"
        ),
        conflict_behavior: UploadConflictBehavior = UploadConflictBehavior.FAIL,
    ) -> dict[str, Any]:
        """Generate Word content without writing locally, then upload it."""

        memory_request = request.model_copy(
            update={"artifact": ArtifactOptions(delivery=ArtifactDelivery.BYTES)}
        )
        artifact = await _document_service().generate_word_document(memory_request)
        item = await _graph_file_service().upload_artifact(
            drive_id,
            destination_path,
            artifact,
            conflict_behavior=conflict_behavior,
        )
        return {
            "artifact": artifact.model_dump(
                mode="json", exclude={"content", "content_base64"}
            ),
            "drive_item": item.model_dump(mode="json", by_alias=True),
        }

    @mcp.tool(
        name="generate_and_upload_powerpoint_presentation",
        description=(
            "Generate a PowerPoint .pptx presentation in memory and upload it "
            "to a OneDrive or SharePoint document-library drive path."
        ),
        tags={"documents", "powerpoint", "files", "sharepoint"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def create_and_upload_powerpoint_presentation(
        request: PowerPointPresentationRequest,
        drive_id: str = Field(..., description="Target OneDrive/SharePoint drive ID"),
        destination_path: str = Field(
            ..., description="Relative destination path including .pptx filename"
        ),
        conflict_behavior: UploadConflictBehavior = UploadConflictBehavior.FAIL,
    ) -> dict[str, Any]:
        """Generate PowerPoint content without writing locally, then upload it."""

        memory_request = request.model_copy(
            update={"artifact": ArtifactOptions(delivery=ArtifactDelivery.BYTES)}
        )
        artifact = await _document_service().generate_powerpoint_presentation(
            memory_request
        )
        item = await _graph_file_service().upload_artifact(
            drive_id,
            destination_path,
            artifact,
            conflict_behavior=conflict_behavior,
        )
        return {
            "artifact": artifact.model_dump(
                mode="json", exclude={"content", "content_base64"}
            ),
            "drive_item": item.model_dump(mode="json", by_alias=True),
        }


def register_power_platform_tools(mcp: FastMCP) -> None:
    """Register supported Dataverse and allowlisted Power Automate tools."""

    @mcp.tool(
        name="get_power_platform_configuration",
        description="Return sanitized Power Platform readiness and named-flow status.",
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": True},
    )
    async def power_platform_configuration() -> dict[str, Any]:
        """Inspect Power Platform readiness without revealing trigger URLs."""

        settings = get_integration_settings().power_platform
        return {
            "configured": settings is not None,
            "environment_host": (
                settings.dataverse_environment_url.host if settings else None
            ),
            "lifecycle_changes_enabled": bool(
                settings and settings.allow_lifecycle_changes
            ),
            "desktop_flow_runs_enabled": bool(
                settings and settings.allow_desktop_flow_runs
            ),
            "desktop_flow_cancellations_enabled": bool(
                settings and settings.allow_desktop_flow_cancellations
            ),
            "named_flows": sorted(settings.named_flows) if settings else [],
            "named_desktop_flows": (
                sorted(settings.named_desktop_flows) if settings else []
            ),
        }

    @mcp.tool(
        name="login_power_platform",
        description=(
            "Explicitly acquire delegated tokens for the configured Dataverse "
            "and named-flow resource audiences. May open the broker/browser."
        ),
        tags={"auth", "power_platform", "power_automate"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def login_power_platform() -> dict[str, Any]:
        """Consent only to administrator-configured non-Graph resource audiences."""

        audiences = _power_platform_audiences()
        provider = MicrosoftAudienceTokenProvider(audiences, allow_interactive=True)
        for audience in sorted(audiences):
            await provider.get_token(audience)
        return {"authenticated": True, "resource_count": len(audiences)}

    @mcp.tool(
        name="list_power_automate_flows",
        description="List solution-aware cloud flows through the Dataverse Web API.",
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": True},
    )
    async def list_power_automate_flows(
        active_only: bool = False,
        top: int = Field(default=100, ge=1, le=5000),
        fetch_all: bool = False,
    ) -> dict[str, Any]:
        """List modern cloud-flow definitions stored in Dataverse."""

        result = await _power_platform_client().list_solution_flows(
            active_only=active_only, top=top, fetch_all=fetch_all
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="get_power_automate_flow",
        description="Get one solution-aware cloud-flow definition from Dataverse.",
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": True},
    )
    async def get_power_automate_flow(workflow_id: UUID) -> dict[str, Any]:
        """Retrieve a cloud-flow workflow row by its immutable ID."""

        result = await _power_platform_client().get_solution_flow(workflow_id)
        return result.model_dump(mode="json")

    @mcp.tool(
        name="activate_power_automate_flow",
        description="Activate a solution-aware cloud flow through Dataverse.",
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def activate_power_automate_flow(
        workflow_id: UUID,
        idempotency_key: str | None = None,
        etag: str = "*",
    ) -> dict[str, Any]:
        """Activate a flow when lifecycle changes are explicitly enabled."""

        result = await _power_platform_client().set_solution_flow_state(
            workflow_id,
            FlowState.ACTIVATED,
            idempotency_key=idempotency_key,
            etag=etag,
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="deactivate_power_automate_flow",
        description="Deactivate a solution-aware cloud flow through Dataverse.",
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def deactivate_power_automate_flow(
        workflow_id: UUID,
        idempotency_key: str | None = None,
        etag: str = "*",
    ) -> dict[str, Any]:
        """Deactivate a flow when lifecycle changes are explicitly enabled."""

        result = await _power_platform_client().set_solution_flow_state(
            workflow_id,
            FlowState.DRAFT,
            idempotency_key=idempotency_key,
            etag=etag,
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="run_power_automate_flow",
        description=(
            "Invoke one named, allowlisted, OAuth-protected Power Automate HTTP "
            "trigger. The caller cannot supply a trigger URL."
        ),
        tags={"power_platform", "power_automate"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def run_power_automate_flow(
        flow_name: str,
        payload: dict[str, Any] | None = None,
        idempotency_key: str | None = None,
        timeout_seconds: float | None = Field(default=None, gt=0, le=300),
    ) -> dict[str, Any]:
        """Run a configured flow by name with bounded timeout and idempotency."""

        result = await _power_platform_client().trigger_solution_flow(
            flow_name,
            payload,
            idempotency_key=idempotency_key,
            timeout_seconds=timeout_seconds,
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="list_power_automate_desktop_flows",
        description=(
            "List published Power Automate desktop flows through the documented "
            "Dataverse workflow API; draft inclusion is explicit."
        ),
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": True},
    )
    async def list_power_automate_desktop_flows(
        top: int = Field(default=100, ge=1, le=5000),
        include_unpublished: bool = False,
        fetch_all: bool = False,
    ) -> dict[str, Any]:
        """List desktop-flow workflow definitions visible to the Dataverse identity."""

        result = await _power_platform_client().list_desktop_flows(
            top=top,
            include_unpublished=include_unpublished,
            fetch_all=fetch_all,
        )
        return result.model_dump(mode="json", by_alias=True)

    @mcp.tool(
        name="get_power_automate_desktop_flow_schema",
        description=(
            "Get the input or output schema for one configured desktop flow. "
            "The caller selects an allowlisted name, never a workflow URL."
        ),
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": True},
    )
    async def get_power_automate_desktop_flow_schema(
        flow_name: str,
        kind: DesktopFlowSchemaKind = DesktopFlowSchemaKind.INPUTS,
    ) -> dict[str, Any]:
        """Read a configured desktop flow's published input or output schema."""

        return await _power_platform_client().get_desktop_flow_schema(flow_name, kind)

    @mcp.tool(
        name="run_power_automate_desktop_flow",
        description=(
            "Run one allowlisted Power Automate desktop flow using its fixed "
            "Dataverse connection or connection reference."
        ),
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def run_power_automate_desktop_flow(
        flow_name: str,
        inputs: dict[str, Any] | None = None,
        run_mode: DesktopFlowRunMode = DesktopFlowRunMode.ATTENDED,
        priority: DesktopFlowPriority = DesktopFlowPriority.NORMAL,
        timeout_seconds: int | None = Field(default=None, ge=1, le=86_400),
    ) -> dict[str, Any]:
        """Queue a configured desktop flow and return its Dataverse session ID."""

        result = await _power_platform_client().run_desktop_flow(
            flow_name,
            inputs,
            run_mode=run_mode,
            priority=priority,
            timeout_seconds=timeout_seconds,
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="get_power_automate_desktop_flow_run",
        description="Get status and timestamps for a Power Automate desktop-flow run.",
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": True},
    )
    async def get_power_automate_desktop_flow_run(
        flow_session_id: UUID,
    ) -> dict[str, Any]:
        """Read documented Dataverse status fields for a desktop-flow session."""

        result = await _power_platform_client().get_desktop_flow_run_status(
            flow_session_id
        )
        return result.model_dump(mode="json", by_alias=True)

    @mcp.tool(
        name="get_power_automate_desktop_flow_outputs",
        description="Read outputs for a completed Power Automate desktop-flow run.",
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": True},
    )
    async def get_power_automate_desktop_flow_outputs(
        flow_session_id: UUID,
    ) -> dict[str, Any]:
        """Read the documented outputs value for a desktop-flow session."""

        return await _power_platform_client().get_desktop_flow_outputs(flow_session_id)

    @mcp.tool(
        name="cancel_power_automate_desktop_flow_run",
        description=(
            "Cancel a queued or running desktop flow through Dataverse when "
            "cancellation is explicitly enabled."
        ),
        tags={"power_platform", "power_automate", "desktop_flows"},
        annotations={"readOnlyHint": False, "destructiveHint": True},
    )
    async def cancel_power_automate_desktop_flow_run(
        flow_session_id: UUID,
    ) -> dict[str, Any]:
        """Cancel a desktop-flow session under both global and integration policy."""

        result = await _power_platform_client().cancel_desktop_flow_run(flow_session_id)
        return result.model_dump(mode="json")


def register_windows_companion_tools(mcp: FastMCP) -> None:
    """Register the authenticated outbound Windows companion control tools."""

    @mcp.tool(
        name="get_windows_companion_configuration",
        description="Return sanitized Windows companion device and action readiness.",
        tags={"windows", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def windows_companion_configuration() -> dict[str, Any]:
        """Inspect configured device aliases without exposing identity material."""

        settings = get_integration_settings().windows_companion
        return {
            "configured": settings is not None,
            "connection_mode": settings.connection_mode if settings else None,
            "devices": {
                alias: {
                    "display_name": device.display_name,
                    "enabled": device.enabled,
                    "allowed_actions": sorted(
                        action.value for action in device.allowed_actions
                    ),
                }
                for alias, device in (settings.devices.items() if settings else [])
            },
        }

    @mcp.tool(
        name="login_windows_companion",
        description=(
            "Explicitly acquire a delegated token for the configured Windows "
            "companion control-plane audience. May open the broker/browser."
        ),
        tags={"auth", "windows", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def login_windows_companion() -> dict[str, Any]:
        """Authenticate only to the configured companion resource audience."""

        settings = get_integration_settings().windows_companion
        if settings is None:
            raise ValueError("Windows companion is not configured.")
        provider = MicrosoftAudienceTokenProvider(
            [settings.token_audience], allow_interactive=True
        )
        await provider.get_token(settings.token_audience)
        return {"authenticated": True}

    @mcp.tool(
        name="get_windows_device_health",
        description="Read authenticated outbound-relay health for an allowlisted laptop.",
        tags={"windows", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def get_windows_device_health(device_alias: str) -> dict[str, Any]:
        """Get connection state and validated companion identity for one device."""

        result = await _windows_companion_client().get_health(device_alias)
        return result.model_dump(mode="json")

    @mcp.tool(
        name="submit_windows_action",
        description=(
            "Submit one typed allowlisted laptop action through the authenticated "
            "outbound relay. Arbitrary shell, process, and URL execution are absent."
        ),
        tags={"windows", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def submit_windows_action(
        device_alias: str,
        action: CompanionAction,
        confirmation: ConfirmationEvidence | None = None,
        idempotency_key: str | None = None,
    ) -> dict[str, Any]:
        """Validate device/action policy and enqueue an expiring action request."""

        integration_settings = get_integration_settings().windows_companion
        if integration_settings is None:
            raise ValueError("Windows companion is not configured.")
        kind = CompanionActionKind(action.kind)
        action_policy = integration_settings.action_policies.get(kind)
        if action_policy is None:
            raise ValueError("Windows action has no configured policy.")
        if action_policy.destructive and not get_settings().allow_destructive_tools:
            raise ValueError(
                "This Windows action requires MICROSOFT_ALLOW_DESTRUCTIVE=true."
            )
        values: dict[str, Any] = {"action": action, "confirmation": confirmation}
        if idempotency_key is not None:
            values["idempotency_key"] = idempotency_key
        request = CompanionActionRequest(**values)
        result = await _windows_companion_client().submit_action(device_alias, request)
        return result.model_dump(mode="json")

    @mcp.tool(
        name="get_windows_action_result",
        description="Read current or final state for a submitted Windows action.",
        tags={"windows", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def get_windows_action_result(
        device_alias: str, action_id: UUID
    ) -> dict[str, Any]:
        """Retrieve a companion action result through the authenticated relay."""

        result = await _windows_companion_client().get_action_result(
            device_alias, action_id
        )
        return result.model_dump(mode="json")


def register_intune_tools(mcp: FastMCP) -> None:
    """Register allowlisted Intune inventory and confirmed remote actions."""

    @mcp.tool(
        name="get_intune_configuration",
        description=(
            "Return sanitized Intune readiness, device count, and stable v1.0 "
            "remote-action capabilities."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def intune_configuration() -> dict[str, Any]:
        """Inspect configured allowlists and action support without credentials."""

        settings = get_integration_settings().intune
        return {
            "configured": settings is not None,
            "allowed_device_count": (
                len(settings.allowed_device_ids) if settings else 0
            ),
            "allowed_actions": (
                sorted(action.value for action in settings.allowed_actions)
                if settings
                else []
            ),
            "tenant_detected_apps_enabled": bool(
                settings and settings.allow_tenant_detected_apps
            ),
            "capabilities": [
                capability.model_dump(mode="json")
                for capability in IntuneService.capabilities()
            ],
        }

    @mcp.tool(
        name="list_intune_managed_devices",
        description="List only managed-device IDs explicitly allowlisted for this agent.",
        tags={"intune", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def list_intune_managed_devices() -> dict[str, Any]:
        """List allowlisted managed-device inventory from Graph v1.0."""

        result = await _intune_service().list_managed_devices()
        return result.model_dump(mode="json", by_alias=True)

    @mcp.tool(
        name="get_intune_managed_device",
        description="Get inventory and compliance data for one allowlisted Intune device.",
        tags={"intune", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def get_intune_managed_device(device_id: UUID) -> dict[str, Any]:
        """Retrieve a managed device after enforcing its UUID allowlist."""

        result = await _intune_service().get_managed_device(device_id)
        return result.model_dump(mode="json", by_alias=True)

    @mcp.tool(
        name="list_intune_detected_apps",
        description=(
            "List tenant-wide Intune detected applications only when that "
            "sensitive inventory capability is explicitly enabled."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": True},
    )
    async def list_intune_detected_apps() -> dict[str, Any]:
        """List detected applications with an independent tenant-wide opt-in."""

        result = await _intune_service().list_detected_apps()
        return result.model_dump(mode="json", by_alias=True)

    @mcp.tool(
        name="sync_intune_device",
        description=(
            "Request an Intune sync for an allowlisted device. Time-bound "
            "confirmation evidence bound to the action and device is required."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def sync_intune_device(
        device_id: UUID, confirmation: IntuneConfirmationEvidence
    ) -> dict[str, Any]:
        """Submit a confirmed syncDevice action through Graph v1.0."""

        result = await _intune_service().sync_device(
            device_id, confirmation=confirmation
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="reboot_intune_device",
        description=(
            "Immediately reboot an allowlisted Intune device after destructive "
            "action acknowledgement and time-bound confirmation."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": True},
    )
    async def reboot_intune_device(
        device_id: UUID, confirmation: IntuneConfirmationEvidence
    ) -> dict[str, Any]:
        """Submit a confirmed and explicitly acknowledged rebootNow action."""

        result = await _intune_service().reboot_now(
            device_id, confirmation=confirmation
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="remote_lock_intune_device",
        description=(
            "Remotely lock an allowlisted Intune device after time-bound "
            "confirmation bound to that device and action."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def remote_lock_intune_device(
        device_id: UUID, confirmation: IntuneConfirmationEvidence
    ) -> dict[str, Any]:
        """Submit a confirmed remoteLock action through Graph v1.0."""

        result = await _intune_service().remote_lock(
            device_id, confirmation=confirmation
        )
        return result.model_dump(mode="json")

    @mcp.tool(
        name="shut_down_intune_device",
        description=(
            "Immediately shut down an allowlisted Intune device after destructive "
            "action acknowledgement and time-bound confirmation."
        ),
        tags={"intune", "devices"},
        annotations={"readOnlyHint": False, "destructiveHint": True},
    )
    async def shut_down_intune_device(
        device_id: UUID, confirmation: IntuneConfirmationEvidence
    ) -> dict[str, Any]:
        """Submit a confirmed and explicitly acknowledged shutDown action."""

        result = await _intune_service().shut_down(device_id, confirmation=confirmation)
        return result.model_dump(mode="json")

    @mcp.tool(
        name="scan_intune_device_with_defender",
        description=(
            "Request a quick or full Microsoft Defender scan on an allowlisted "
            "Intune device with time-bound confirmation."
        ),
        tags={"intune", "devices", "security"},
        annotations={"readOnlyHint": False, "destructiveHint": False},
    )
    async def scan_intune_device_with_defender(
        device_id: UUID,
        quick_scan: bool,
        confirmation: IntuneConfirmationEvidence,
    ) -> dict[str, Any]:
        """Submit a confirmed windowsDefenderScan action through Graph v1.0."""

        result = await _intune_service().windows_defender_scan(
            device_id,
            quick_scan=quick_scan,
            confirmation=confirmation,
        )
        return result.model_dump(mode="json")
