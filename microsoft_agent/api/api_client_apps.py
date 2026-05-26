import os
import sys
from typing import Any

from microsoft_agent.auth import AuthManager

CLIENT_ID = os.environ.get("OIDC_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite",
    "Tasks.ReadWrite",
    "Contacts.ReadWrite",
    "Group.ReadWrite.All",
    "Directory.Read.All",
    "Sites.Read.All",
    "Chat.Read",
    "ChatMessage.Read.All",
    "ChannelMessage.Read.All",
    "ServiceHealth.Read.All",
    "ServiceMessage.Read.All",
    "Domain.ReadWrite.All",
    "Organization.ReadWrite.All",
    "OnlineMeetings.ReadWrite",
    "CallRecords.Read.All",
    "Presence.Read.All",
    "User.Invite.All",
    "SecurityEvents.ReadWrite.All",
    "SecurityIncident.ReadWrite.All",
    "ThreatHunting.Read.All",
    "AuditLog.Read.All",
    "Reports.Read.All",
    "Application.ReadWrite.All",
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "IdentityRiskEvent.Read.All",
    "IdentityRiskyUser.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory",
    "EntitlementManagement.Read.All",
    "AccessReview.Read.All",
    "LifecycleWorkflows.Read.All",
]

# Only create global auth_manager if not in test mode
auth_manager: AuthManager | None
if not os.environ.get("TESTING"):
    auth_manager = AuthManager(CLIENT_ID, AUTHORITY, SCOPES)
else:
    auth_manager = None

from microsoft_agent.api.api_client_base import MicrosoftGraphApiBase


class MicrosoftGraphApiApps(MicrosoftGraphApiBase):
    async def get_excel_workbook(
        self, drive_id: str, item_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Excel workbook."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel workbook: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_excel_worksheets(
        self, drive_id: str, item_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List Excel worksheets."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.worksheets.worksheets_request_builder import (
            WorksheetsRequestBuilder,
        )

        try:
            query_params = (
                WorksheetsRequestBuilder.WorksheetsRequestBuilderGetQueryParameters()
            )
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = WorksheetsRequestBuilder.WorksheetsRequestBuilderGetQueryParameters().WorksheetsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing excel worksheets: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_excel_worksheet(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Get Excel worksheet."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel worksheet: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_excel_tables(
        self, drive_id: str, item_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List Excel tables."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.tables.tables_request_builder import (
            TablesRequestBuilder,
        )

        try:
            query_params = TablesRequestBuilder.TablesRequestBuilderGetQueryParameters()
            if params:
                if "$select" in params:
                    query_params.select = params["$select"].split(",")

            request_config = (
                TablesRequestBuilder.TablesRequestBuilderGetRequestConfiguration(
                    query_parameters=query_params,
                    options=[ResponseHandlerOption(NativeResponseHandler())],
                )
            )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing excel tables: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_excel_table(
        self, drive_id: str, item_id: str, table_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Excel table."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.by_workbook_table_id(table_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.tables.by_workbook_table_id(table_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting excel table: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_onenote_notebook_sections(
        self, notebook_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List Onenote notebook sections."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.onenote.notebooks.by_notebook_id(
                notebook_id
            ).sections.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.onenote.notebooks.by_notebook_id(
                notebook_id
            ).sections.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing onenote sections: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_onenote_section_pages(
        self, onenoteSection_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List Onenote section pages."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.onenote.sections.by_onenote_section_id(
                onenoteSection_id
            ).pages.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.onenote.sections.by_onenote_section_id(
                    onenoteSection_id
                ).pages.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing onenote pages: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_onenote_page_content(
        self, onenotePage_id: str, params: dict | None = None
    ) -> Any:
        """Get Onenote page content."""
        try:
            response = await self.client.me.onenote.pages.by_onenote_page_id(
                onenotePage_id
            ).content.get()
            if isinstance(response, bytes):
                return {"content": response.decode("utf-8")}
            return {"error": "Unexpected response type"}
        except Exception as e:
            print(f"Error getting onenote page content: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_onenote_page(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Onenote page."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            html_content = data.get("content", "")

            request_config = (
                self.client.me.onenote.pages.to_post_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.onenote.pages.post(
                html_content.encode("utf-8"), request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating onenote page: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_todo_task_lists(self, params: dict | None = None) -> dict[str, Any]:
        """List Todo task lists."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.todo.lists.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing todo task lists: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_todo_tasks(
        self, todoTaskList_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List Todo tasks."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing todo tasks: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_todo_task(
        self, todoTaskList_id: str, todoTask_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = (
                self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .to_get_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting todo task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_todo_task(
        self, todoTaskList_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.todo_task import TodoTask

        try:
            task = TodoTask()
            task.title = data.get("title")

            request_config = self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.todo.lists.by_todo_task_list_id(
                todoTaskList_id
            ).tasks.post(task, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating todo task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_todo_task(
        self,
        todoTaskList_id: str,
        todoTask_id: str,
        data: dict[str, Any],
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Update Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.todo_task import TodoTask

        try:
            task = TodoTask()
            task.title = data.get("title")

            request_config = (
                self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .to_patch_request_configuration()
            )
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = (
                await self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .patch(task, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating todo task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def delete_todo_task(
        self, todoTaskList_id: str, todoTask_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Delete Todo task."""
        try:
            await (
                self.client.me.todo.lists.by_todo_task_list_id(todoTaskList_id)
                .tasks.by_todo_task_id(todoTask_id)
                .delete()
            )
            return {"status": "success"}
        except Exception as e:
            print(f"Error deleting todo task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_planner_tasks(self, params: dict | None = None) -> dict[str, Any]:
        """List Planner tasks."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.me.planner.tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.me.planner.tasks.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing planner tasks: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_planner_plan(
        self, plannerPlan_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Planner plan."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting planner plan: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def list_plan_tasks(
        self, plannerPlan_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """List tasks for a Planner plan."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).tasks.to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.plans.by_planner_plan_id(
                plannerPlan_id
            ).tasks.get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error listing plan tasks: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def get_planner_task(
        self, plannerTask_id: str, params: dict | None = None
    ) -> dict[str, Any]:
        """Get Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        try:
            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).to_get_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).get(request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error getting planner task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def create_planner_task(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            task = PlannerTask()
            task.title = data.get("title")
            task.plan_id = data.get("planId")

            request_config = self.client.planner.tasks.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.post(
                task, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error creating planner task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_planner_task(
        self, plannerTask_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            task = PlannerTask()
            task.title = data.get("title")

            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).patch(task, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating planner task: {e}", file=sys.stderr)
            return {"error": str(e)}

    async def update_planner_task_details(
        self, plannerTask_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Update Planner task details."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task_details import PlannerTaskDetails

        try:
            details = PlannerTaskDetails()
            details.description = data.get("description")

            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.to_patch_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.patch(details, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Error updating planner task details: {e}", file=sys.stderr)
            return {"error": str(e)}
