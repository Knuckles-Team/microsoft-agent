from typing import Any

from microsoft_agent.api._graph_models import (
    graph_model_from_dict,
    validated_planner_etag,
)
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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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

            request_config = WorksheetsRequestBuilder.WorksheetsRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
            if params and params.get("Workbook-Session-Id"):
                request_config.headers.add(
                    "Workbook-Session-Id", params["Workbook-Session-Id"]
                )
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.get(request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_onenote_page(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create a OneNote page from the raw HTML body required by Graph."""
        from kiota_abstractions.method import Method
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_abstractions.request_information import RequestInformation
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.onenote_page import OnenotePage

        try:
            html_content = data.get("content", "")
            if not isinstance(html_content, str) or not html_content.strip():
                return {"error": "OneNote page HTML content is required"}

            pages = self.client.me.onenote.pages
            request_config = pages.to_post_request_configuration()
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            # The generated SDK models this endpoint as JSON, but Graph's OneNote
            # create-page contract requires a raw text/html (or multipart) body.
            request_info = RequestInformation(
                Method.POST, pages.url_template, pages.path_parameters
            )
            request_info.configure(request_config)
            request_info.headers.try_add("Accept", "application/json")
            request_info.set_stream_content(html_content.encode("utf-8"), "text/html")
            native_response = await pages.request_adapter.send_async(
                request_info, OnenotePage, {}
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_todo_task(
        self, todoTaskList_id: str, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Todo task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.todo_task import TodoTask

        try:
            task = graph_model_from_dict(data, TodoTask)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            task = graph_model_from_dict(data, TodoTask)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_planner_task(
        self, data: dict[str, Any], params: dict | None = None
    ) -> dict[str, Any]:
        """Create Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            task = graph_model_from_dict(data, PlannerTask)

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
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_planner_task(
        self,
        plannerTask_id: str,
        data: dict[str, Any],
        params: dict | None = None,
        etag: str | None = None,
    ) -> dict[str, Any]:
        """Update Planner task."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task import PlannerTask

        try:
            if_match = validated_planner_etag(etag, params)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
            task = graph_model_from_dict(data, PlannerTask)
            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).to_patch_request_configuration()
            request_config.headers.add("If-Match", if_match)
            request_config.headers.add("Prefer", "return=representation")
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).patch(task, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def update_planner_task_details(
        self,
        plannerTask_id: str,
        data: dict[str, Any],
        params: dict | None = None,
        etag: str | None = None,
    ) -> dict[str, Any]:
        """Update Planner task details."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.models.planner_task_details import PlannerTaskDetails

        try:
            if_match = validated_planner_etag(etag, params)
        except ValueError as exc:
            return {"error": str(exc)}

        try:
            details = graph_model_from_dict(data, PlannerTaskDetails)
            request_config = self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.to_patch_request_configuration()
            request_config.headers.add("If-Match", if_match)
            request_config.headers.add("Prefer", "return=representation")
            request_config.options.append(
                ResponseHandlerOption(NativeResponseHandler())
            )

            native_response = await self.client.planner.tasks.by_planner_task_id(
                plannerTask_id
            ).details.patch(details, request_configuration=request_config)
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def create_excel_chart(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        data: dict[str, Any] | None = None,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Create a chart in an Excel worksheet."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.worksheets.item.charts.add.add_post_request_body import (
            AddPostRequestBody,
        )
        from msgraph.generated.drives.item.items.item.workbook.worksheets.item.charts.add.add_request_builder import (
            AddRequestBuilder,
        )

        chart_data = data or {}
        request_body = AddPostRequestBody(
            type=chart_data.get("type"),
            series_by=chart_data.get("seriesBy"),
            additional_data={
                key: value
                for key, value in chart_data.items()
                if key not in {"type", "seriesBy"}
            },
        )
        request_config = AddRequestBuilder.AddRequestBuilderPostRequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())]
        )
        if params and params.get("Workbook-Session-Id"):
            request_config.headers.add(
                "Workbook-Session-Id", params["Workbook-Session-Id"]
            )

        try:
            native_response = (
                await self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .charts.add.post(request_body, request_configuration=request_config)
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def format_excel_range(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        address: str,
        data: dict[str, Any] | None = None,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Update formatting for an addressed Excel worksheet range."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.drives.item.items.item.workbook.worksheets.item.range_with_address.format.format_request_builder import (
            FormatRequestBuilder,
        )
        from msgraph.generated.models.workbook_range_format import WorkbookRangeFormat

        format_data = data or {}
        mapped_keys = {
            "columnWidth",
            "horizontalAlignment",
            "rowHeight",
            "verticalAlignment",
            "wrapText",
        }
        request_body = WorkbookRangeFormat(
            column_width=format_data.get("columnWidth"),
            horizontal_alignment=format_data.get("horizontalAlignment"),
            row_height=format_data.get("rowHeight"),
            vertical_alignment=format_data.get("verticalAlignment"),
            wrap_text=format_data.get("wrapText"),
            additional_data={
                key: value
                for key, value in format_data.items()
                if key not in mapped_keys
            },
        )
        request_config = (
            FormatRequestBuilder.FormatRequestBuilderPatchRequestConfiguration(
                options=[ResponseHandlerOption(NativeResponseHandler())]
            )
        )
        if params and params.get("Workbook-Session-Id"):
            request_config.headers.add(
                "Workbook-Session-Id", params["Workbook-Session-Id"]
            )

        try:
            range_builder = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .range_with_address(address)
            )
            native_response = await range_builder.format.patch(
                request_body, request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def get_excel_range(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        address: str,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Get an addressed range from an Excel worksheet."""
        from kiota_abstractions.base_request_configuration import RequestConfiguration
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption

        request_config = RequestConfiguration(
            options=[ResponseHandlerOption(NativeResponseHandler())]
        )
        if params and params.get("Workbook-Session-Id"):
            request_config.headers.add(
                "Workbook-Session-Id", params["Workbook-Session-Id"]
            )

        try:
            range_builder = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .range_with_address(address)
            )
            native_response = await range_builder.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def sort_excel_range(
        self,
        drive_id: str,
        item_id: str,
        worksheet_id: str,
        address: str,
        data: dict[str, Any] | None = None,
        params: dict | None = None,
    ) -> dict[str, Any]:
        """Apply a sort operation to an addressed Excel worksheet range."""
        from kiota_abstractions.base_request_configuration import RequestConfiguration
        from kiota_abstractions.method import Method
        from kiota_abstractions.request_information import RequestInformation
        from msgraph.generated.drives.item.items.item.workbook.worksheets.item.tables.item.sort.apply.apply_post_request_body import (
            ApplyPostRequestBody,
        )
        from msgraph.generated.models.o_data_errors.o_data_error import ODataError
        from msgraph.generated.models.workbook_icon import WorkbookIcon
        from msgraph.generated.models.workbook_sort_field import WorkbookSortField

        sort_data = data or {}
        fields = []
        for field_data in sort_data.get("fields", []):
            icon_data = field_data.get("icon")
            icon = None
            if icon_data:
                icon = WorkbookIcon(
                    set=icon_data.get("set"),
                    index=icon_data.get("index"),
                    additional_data={
                        key: value
                        for key, value in icon_data.items()
                        if key not in {"set", "index"}
                    },
                )
            fields.append(
                WorkbookSortField(
                    key=field_data.get("key"),
                    ascending=field_data.get("ascending"),
                    color=field_data.get("color"),
                    data_option=field_data.get("dataOption"),
                    sort_on=field_data.get("sortOn"),
                    icon=icon,
                    additional_data={
                        key: value
                        for key, value in field_data.items()
                        if key
                        not in {
                            "key",
                            "ascending",
                            "color",
                            "dataOption",
                            "sortOn",
                            "icon",
                        }
                    },
                )
            )

        request_body = ApplyPostRequestBody(
            fields=fields,
            match_case=sort_data.get("matchCase"),
            method=sort_data.get("method"),
            additional_data={
                key: value
                for key, value in sort_data.items()
                if key not in {"fields", "matchCase", "method"}
            },
        )

        try:
            range_builder = (
                self.client.drives.by_drive_id(drive_id)
                .items.by_drive_item_id(item_id)
                .workbook.worksheets.by_workbook_worksheet_id(worksheet_id)
                .range_with_address(address)
            )
            sort_builder = range_builder.sort
            url_template = f"{sort_builder.url_template.partition('{?')[0]}/apply"
            request_info = RequestInformation(
                Method.POST, url_template, sort_builder.path_parameters
            )
            request_info.configure(RequestConfiguration())
            request_info.headers.try_add("Accept", "application/json")
            if params and params.get("Workbook-Session-Id"):
                request_info.headers.try_add(
                    "Workbook-Session-Id", params["Workbook-Session-Id"]
                )
            request_info.set_content_from_parsable(
                sort_builder.request_adapter, "application/json", request_body
            )
            await sort_builder.request_adapter.send_no_response_content_async(
                request_info, {"4XX": ODataError, "5XX": ODataError}
            )
            return {"status": "success"}
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}

    async def list_onenote_notebooks(
        self, params: dict | None = None
    ) -> dict[str, Any]:
        """List notebooks owned by the current user."""
        from kiota_abstractions.native_response_handler import NativeResponseHandler
        from kiota_http.middleware.options import ResponseHandlerOption
        from msgraph.generated.users.item.onenote.notebooks.notebooks_request_builder import (
            NotebooksRequestBuilder,
        )

        query_params = (
            NotebooksRequestBuilder.NotebooksRequestBuilderGetQueryParameters()
        )
        if params:
            if "$select" in params:
                query_params.select = params["$select"].split(",")
            if "$expand" in params:
                query_params.expand = params["$expand"].split(",")
            if "$filter" in params:
                query_params.filter = params["$filter"]
            if "$orderby" in params:
                query_params.orderby = params["$orderby"].split(",")
            if "$search" in params:
                query_params.search = params["$search"]
            if "$skip" in params:
                query_params.skip = int(params["$skip"])
            if "$top" in params:
                query_params.top = int(params["$top"])
            if "$count" in params:
                query_params.count = str(params["$count"]).lower() == "true"

        request_config = (
            NotebooksRequestBuilder.NotebooksRequestBuilderGetRequestConfiguration(
                query_parameters=query_params,
                options=[ResponseHandlerOption(NativeResponseHandler())],
            )
        )

        try:
            native_response = await self.client.me.onenote.notebooks.get(
                request_configuration=request_config
            )
            native_response.raise_for_status()
            return native_response.json()
        except Exception as e:
            print(f"Operation failed: {type(e).__name__}")
            return {"error": "Operation failed"}
