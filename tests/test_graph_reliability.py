"""Focused regression tests for Microsoft Graph tool and wrapper boundaries."""

from unittest.mock import AsyncMock, MagicMock

import pytest
from fastmcp import FastMCP

from microsoft_agent.api_client import MicrosoftGraphApi


def _api_with_client(client: MagicMock) -> MicrosoftGraphApi:
    api = object.__new__(MicrosoftGraphApi)
    api.client = client
    return api


def _native_response(payload: dict | None = None) -> MagicMock:
    response = MagicMock()
    response.raise_for_status = MagicMock()
    response.json.return_value = payload or {"value": []}
    return response


@pytest.mark.asyncio
async def test_get_me_accepts_and_applies_query_params() -> None:
    from msgraph.generated.users.item.user_item_request_builder import (
        UserItemRequestBuilder,
    )

    response = _native_response({"id": "user-1"})
    client = MagicMock()
    client.me.get = AsyncMock(return_value=response)
    api = _api_with_client(client)

    result = await api.get_me(
        params={"$select": "id,displayName", "$expand": "manager"}
    )

    assert result == {"id": "user-1"}
    awaited = client.me.get.await_args
    assert awaited is not None
    request_config = awaited.kwargs["request_configuration"]
    assert isinstance(
        request_config,
        UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration,
    )
    assert request_config.query_parameters.select == ["id", "displayName"]
    assert request_config.query_parameters.expand == ["manager"]


@pytest.mark.asyncio
async def test_sharepoint_wrappers_call_generated_request_builders() -> None:
    from msgraph.generated.sites.item.drives.item.drive_item_request_builder import (
        DriveItemRequestBuilder,
    )
    from msgraph.generated.sites.item.items.item.base_item_item_request_builder import (
        BaseItemItemRequestBuilder,
    )
    from msgraph.generated.sites.item.items.items_request_builder import (
        ItemsRequestBuilder,
    )

    response = _native_response()
    client = MagicMock()
    site = client.sites.by_site_id.return_value
    site.drives.by_drive_id.return_value.get = AsyncMock(return_value=response)
    site.items.get = AsyncMock(return_value=response)
    site.items.by_base_item_id.return_value.get = AsyncMock(return_value=response)
    api = _api_with_client(client)

    assert await api.get_site_drive_by_id(
        "site-1", "drive-1", {"$select": "id,name"}
    ) == {"value": []}
    assert await api.list_site_items("site-1", {"$top": "5", "$count": "true"}) == {
        "value": []
    }
    assert await api.get_site_item(
        "site-1", "item-1", {"$expand": "parentReference"}
    ) == {"value": []}

    drive_awaited = site.drives.by_drive_id.return_value.get.await_args
    items_awaited = site.items.get.await_args
    item_awaited = site.items.by_base_item_id.return_value.get.await_args
    assert drive_awaited is not None
    assert items_awaited is not None
    assert item_awaited is not None
    drive_config = drive_awaited.kwargs["request_configuration"]
    items_config = items_awaited.kwargs["request_configuration"]
    item_config = item_awaited.kwargs["request_configuration"]
    assert isinstance(
        drive_config,
        DriveItemRequestBuilder.DriveItemRequestBuilderGetRequestConfiguration,
    )
    assert drive_config.query_parameters.select == ["id", "name"]
    assert isinstance(
        items_config, ItemsRequestBuilder.ItemsRequestBuilderGetRequestConfiguration
    )
    assert items_config.query_parameters.top == 5
    assert items_config.query_parameters.count is True
    assert isinstance(
        item_config,
        BaseItemItemRequestBuilder.BaseItemItemRequestBuilderGetRequestConfiguration,
    )
    assert item_config.query_parameters.expand == ["parentReference"]


@pytest.mark.asyncio
async def test_onenote_notebooks_wrapper_applies_collection_params() -> None:
    from msgraph.generated.users.item.onenote.notebooks.notebooks_request_builder import (
        NotebooksRequestBuilder,
    )

    response = _native_response()
    client = MagicMock()
    client.me.onenote.notebooks.get = AsyncMock(return_value=response)
    api = _api_with_client(client)

    result = await api.list_onenote_notebooks(
        {"$select": "id,displayName", "$top": "10"}
    )

    assert result == {"value": []}
    awaited = client.me.onenote.notebooks.get.await_args
    assert awaited is not None
    request_config = awaited.kwargs["request_configuration"]
    assert isinstance(
        request_config,
        NotebooksRequestBuilder.NotebooksRequestBuilderGetRequestConfiguration,
    )
    assert request_config.query_parameters.select == ["id", "displayName"]
    assert request_config.query_parameters.top == 10


@pytest.mark.asyncio
async def test_excel_chart_format_and_range_use_the_addressed_worksheet() -> None:
    from msgraph.generated.drives.item.items.item.workbook.worksheets.item.charts.add.add_post_request_body import (
        AddPostRequestBody,
    )
    from msgraph.generated.models.workbook_range_format import WorkbookRangeFormat

    response = _native_response({"id": "result-1"})
    client = MagicMock()
    worksheet = client.drives.by_drive_id.return_value.items.by_drive_item_id.return_value.workbook.worksheets.by_workbook_worksheet_id.return_value
    worksheet.charts.add.post = AsyncMock(return_value=response)
    addressed_range = worksheet.range_with_address.return_value
    addressed_range.format.patch = AsyncMock(return_value=response)
    addressed_range.get = AsyncMock(return_value=response)
    api = _api_with_client(client)

    chart_result = await api.create_excel_chart(
        "drive-1",
        "item-1",
        "sheet-1",
        {"type": "ColumnStacked", "sourceData": "A1:B2", "seriesBy": "Auto"},
    )
    format_result = await api.format_excel_range(
        "drive-1",
        "item-1",
        "sheet-1",
        "A1:B2",
        {"columnWidth": 135, "wrapText": False},
    )
    range_result = await api.get_excel_range("drive-1", "item-1", "sheet-1", "A1:B2")

    assert chart_result == {"id": "result-1"}
    assert format_result == {"id": "result-1"}
    assert range_result == {"id": "result-1"}
    chart_awaited = worksheet.charts.add.post.await_args
    assert chart_awaited is not None
    chart_body = chart_awaited.args[0]
    assert isinstance(chart_body, AddPostRequestBody)
    assert chart_body.type == "ColumnStacked"
    assert chart_body.series_by == "Auto"
    assert chart_body.additional_data["sourceData"] == "A1:B2"
    worksheet.range_with_address.assert_any_call("A1:B2")
    format_awaited = addressed_range.format.patch.await_args
    assert format_awaited is not None
    format_body = format_awaited.args[0]
    assert isinstance(format_body, WorkbookRangeFormat)
    assert format_body.column_width == 135
    assert format_body.wrap_text is False


@pytest.mark.asyncio
async def test_sort_excel_range_posts_to_apply_endpoint() -> None:
    response_adapter = MagicMock()
    response_adapter.send_no_response_content_async = AsyncMock(return_value=None)
    writer = response_adapter.get_serialization_writer_factory.return_value.get_serialization_writer.return_value
    writer.get_serialized_content.return_value = b"{}"

    client = MagicMock()
    worksheet = client.drives.by_drive_id.return_value.items.by_drive_item_id.return_value.workbook.worksheets.by_workbook_worksheet_id.return_value
    sort_builder = worksheet.range_with_address.return_value.sort
    sort_builder.url_template = (
        "{+baseurl}/drives/{drive%2Did}/items/{driveItem%2Did}/workbook/"
        "worksheets/{workbookWorksheet%2Did}/range(address='{address}')/sort"
        "{?%24expand,%24select}"
    )
    sort_builder.path_parameters = {
        "drive%2Did": "drive-1",
        "driveItem%2Did": "item-1",
        "workbookWorksheet%2Did": "sheet-1",
        "address": "A1:C10",
    }
    sort_builder.request_adapter = response_adapter
    api = _api_with_client(client)

    result = await api.sort_excel_range(
        "drive-1",
        "item-1",
        "sheet-1",
        "A1:C10",
        {
            "fields": [{"key": 0, "ascending": True}],
            "hasHeaders": True,
            "orientation": "Rows",
        },
        {"Workbook-Session-Id": "session-1"},
    )

    assert result == {"status": "success"}
    send_awaited = response_adapter.send_no_response_content_async.await_args
    assert send_awaited is not None
    request_info = send_awaited.args[0]
    assert request_info.url_template.endswith("/sort/apply")
    assert request_info.headers.get("Workbook-Session-Id") == {"session-1"}
    request_body = writer.write_object_value.call_args.args[1]
    assert request_body.fields[0].key == 0
    assert request_body.fields[0].ascending is True
    assert request_body.additional_data == {
        "hasHeaders": True,
        "orientation": "Rows",
    }


@pytest.mark.asyncio
async def test_list_excel_worksheets_uses_valid_request_configuration() -> None:
    from msgraph.generated.drives.item.items.item.workbook.worksheets.worksheets_request_builder import (
        WorksheetsRequestBuilder,
    )

    response = _native_response()
    client = MagicMock()
    worksheets = client.drives.by_drive_id.return_value.items.by_drive_item_id.return_value.workbook.worksheets
    worksheets.get = AsyncMock(return_value=response)
    api = _api_with_client(client)

    result = await api.list_excel_worksheets(
        "drive-1", "item-1", {"$select": "id,name"}
    )

    assert result == {"value": []}
    awaited = worksheets.get.await_args
    assert awaited is not None
    request_config = awaited.kwargs["request_configuration"]
    assert isinstance(
        request_config,
        WorksheetsRequestBuilder.WorksheetsRequestBuilderGetRequestConfiguration,
    )
    assert request_config.query_parameters.select == ["id", "name"]


@pytest.mark.asyncio
async def test_files_action_tool_awaits_current_client_methods() -> None:
    from microsoft_agent.mcp_server import register_files_tools

    graph_api = MagicMock()
    graph_api.list_users = AsyncMock(return_value={"value": []})
    graph_api.list_excel_worksheets = AsyncMock(return_value={"value": []})
    graph_api.format_excel_range = AsyncMock(return_value={"status": "success"})
    graph_api.sort_excel_range = AsyncMock(return_value={"status": "success"})
    mcp = FastMCP("graph-reliability-test")
    register_files_tools(mcp)
    files = await mcp.get_tool("microsoft_files")

    await files.fn(
        action="list_users",
        params_json='{"params":{"$top":"1"}}',
        client=graph_api,
        ctx=None,
    )
    await files.fn(
        action="list_excel_worksheets",
        params_json='{"drive_id":"drive-1","item_id":"item-1","params":null}',
        client=graph_api,
        ctx=None,
    )
    await files.fn(
        action="format_excel_range",
        params_json=(
            '{"drive_id":"drive-1","item_id":"item-1",'
            '"worksheet_id":"sheet-1","address":"A1:B2",'
            '"data":{"wrapText":true},"params":null}'
        ),
        client=graph_api,
        ctx=None,
    )
    await files.fn(
        action="sort_excel_range",
        params_json=(
            '{"drive_id":"drive-1","item_id":"item-1",'
            '"worksheet_id":"sheet-1","address":"A1:B2",'
            '"data":{"fields":[{"key":0}]},"params":null}'
        ),
        client=graph_api,
        ctx=None,
    )

    graph_api.list_users.assert_awaited_once_with(params={"$top": "1"})
    graph_api.list_excel_worksheets.assert_awaited_once_with(
        drive_id="drive-1", item_id="item-1"
    )
    graph_api.format_excel_range.assert_awaited_once_with(
        drive_id="drive-1",
        worksheet_id="sheet-1",
        item_id="item-1",
        address="A1:B2",
        data={"wrapText": True},
    )
    graph_api.sort_excel_range.assert_awaited_once_with(
        drive_id="drive-1",
        item_id="item-1",
        worksheet_id="sheet-1",
        address="A1:B2",
        data={"fields": [{"key": 0}]},
    )
