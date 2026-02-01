#!/usr/bin/env python
# coding: utf-8

import requests
from typing import Dict, Optional, Any
from urllib.parse import urljoin


class Api:
    def __init__(
        self,
        base_url: str = "https://graph.microsoft.com/v1.0",
        token: Optional[str] = None,
    ):
        self.base_url = base_url
        self.token = token
        self._session = requests.Session()

    def get_headers(self) -> Dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        return headers

    def request(
        self, method: str, endpoint: str, data: Dict = None, params: Dict = None
    ) -> Any:
        url = (
            urljoin(self.base_url, endpoint.lstrip("/"))
            if not endpoint.startswith("http")
            else endpoint
        )
        headers = self.get_headers()
        response = self._session.request(
            method=method, url=url, headers=headers, json=data, params=params
        )
        if response.status_code >= 400:
            try:
                err_msg = response.json()
            except Exception:
                err_msg = response.text
            raise Exception(f"Error {response.status_code}: {err_msg}")
        if response.status_code == 204:
            return {"status": "success"}
        try:
            return response.json()
        except Exception:
            return response.text

    # --- Auto-generated Methods (183 endpoints) ---

    def getmemberobjects_directoryobject(
        self, directoryObject_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """directoryObject: getMemberObjects"""
        endpoint = f"/directoryObjects/{directoryObject_id}/getMemberObjects"
        return self.request("POST", endpoint, data=None, params=params)

    def list_members_group(self, group_id: str, params: Dict = None) -> Any:
        """List group members"""
        endpoint = f"/groups/{group_id}/members"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_members_group(
        self, group_id: str, member_id: str, params: Dict = None
    ) -> Any:
        """Remove member"""
        endpoint = f"/groups/{group_id}/members/{member_id}/$ref"
        return self.request("DELETE", endpoint, data=None, params=params)

    def list_owners_group(self, group_id: str, params: Dict = None) -> Any:
        """List group owners"""
        endpoint = f"/groups/{group_id}/owners"
        return self.request("GET", endpoint, data=None, params=params)

    def overview_resources_planner(self, group_id: str, params: Dict = None) -> Any:
        """Use the Planner REST API"""
        endpoint = f"/groups/{group_id}/planner/plans"
        return self.request("GET", endpoint, data=None, params=params)

    def forward_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: forward"""
        endpoint = f"/me/events/{event_id}/forward"
        return self.request("POST", endpoint, data=None, params=params)

    def get_conversation_group(
        self, group_id: str, conversation_id: str, params: Dict = None
    ) -> Any:
        """Get conversation"""
        endpoint = f"/groups/{group_id}/conversations/{conversation_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_driveitem(self, drive_id: str, item_id: str, params: Dict = None) -> Any:
        """Delete a DriveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def follow_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Follow drive item"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/follow"
        return self.request("POST", endpoint, data=None, params=params)

    def copynotebook_notebook(
        self, notebook_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """notebook: copyNotebook"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}/copyNotebook"
        return self.request("POST", endpoint, data=None, params=params)

    def snoozereminder_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: snoozeReminder"""
        endpoint = f"/me/events/{event_id}/snoozeReminder"
        return self.request("POST", endpoint, data=None, params=params)

    def getoffice365groupsactivitygroupcounts_reportroot(
        self, period_value: str, params: Dict = None
    ) -> Any:
        """reportRoot: getOffice365GroupsActivityGroupCounts"""
        endpoint = (
            f"/reports/getOffice365GroupsActivityGroupCounts(period='{period_value}')"
        )
        return self.request("GET", endpoint, data=None, params=params)

    def get_calendar(self, id_or_userPrincipalName: str, params: Dict = None) -> Any:
        """Get calendar"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendar"
        return self.request("GET", endpoint, data=None, params=params)

    def list_transitivememberof_group(self, group_id: str, params: Dict = None) -> Any:
        """List group transitive memberOf"""
        endpoint = f"/groups/{group_id}/transitiveMemberOf"
        return self.request("GET", endpoint, data=None, params=params)

    def post_rejectedsenders_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create rejectedSender"""
        endpoint = f"/groups/{group_id}/rejectedSenders/$ref"
        return self.request("POST", endpoint, data=None, params=params)

    def permanentdelete_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: permanentDelete"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/permanentDelete"
        return self.request("POST", endpoint, data=None, params=params)

    def copytonotebook_section(
        self, section_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """section: copyToNotebook"""
        endpoint = f"/me/onenote/sections/{section_id}/copyToNotebook"
        return self.request("POST", endpoint, data=None, params=params)

    def list_user(self, params: Dict = None) -> Any:
        """List users"""
        endpoint = "/users"
        return self.request("GET", endpoint, data=None, params=params)

    def revokesigninsessions_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """user: revokeSignInSessions"""
        endpoint = f"/users/{id_or_userPrincipalName}/revokeSignInSessions"
        return self.request("POST", endpoint, data=None, params=params)

    def list_sections_notebook(self, notebook_id: str, params: Dict = None) -> Any:
        """List sections"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}/sections"
        return self.request("GET", endpoint, data=None, params=params)

    def restore_driveitem(
        self, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: restore"""
        endpoint = f"/me/drive/items/{item_id}/restore"
        return self.request("POST", endpoint, data=None, params=params)

    def preview_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: preview"""
        endpoint = f"/drives/{driveId}/items/{itemId}/preview"
        return self.request("POST", endpoint, data=None, params=params)

    def dismissreminder_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: dismissReminder"""
        endpoint = f"/me/events/{event_id}/dismissReminder"
        return self.request("POST", endpoint, data=None, params=params)

    def move_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Move a driveItem to a new folder"""
        endpoint = f"/drives/{drive_id}/items/{item_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def api_overview_resources_onenote(self, params: Dict = None) -> Any:
        """Use the OneNote REST API"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def accept_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: accept"""
        endpoint = f"/me/events/{event_id}/accept"
        return self.request("POST", endpoint, data=None, params=params)

    def update_chart(
        self,
        item_id: str,
        id_or_name: str,
        name: str,
        data: Dict = None,
        params: Dict = None,
    ) -> Any:
        """Update chart"""
        endpoint = (
            f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}"
        )
        return self.request("PATCH", endpoint, data=None, params=params)

    def delete_rejectedsenders_group(self, group_id: str, params: Dict = None) -> Any:
        """Remove rejectedSender"""
        endpoint = f"/groups/{group_id}/rejectedSenders/$ref"
        return self.request("DELETE", endpoint, data=None, params=params)

    def list_subscription(self, params: Dict = None) -> Any:
        """List subscriptions"""
        endpoint = "/subscriptions"
        return self.request("GET", endpoint, data=None, params=params)

    def add_chartcollection(
        self, item_id: str, id_or_name: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """ChartCollection: add"""
        endpoint = (
            f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/add"
        )
        return self.request("POST", endpoint, data=None, params=params)

    def update_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update DriveItem properties"""
        endpoint = f"/drives/{drive_id}/items/{item_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def list_threads_group(self, group_id: str, params: Dict = None) -> Any:
        """List threads"""
        endpoint = f"/groups/{group_id}/threads"
        return self.request("GET", endpoint, data=None, params=params)

    def list_conversations_group(self, group_id: str, params: Dict = None) -> Any:
        """List conversations"""
        endpoint = f"/groups/{group_id}/conversations"
        return self.request("GET", endpoint, data=None, params=params)

    def list_cloudpcs_user(self, params: Dict = None) -> Any:
        """List cloudPCs for user"""
        endpoint = "/me/cloudPCs"
        return self.request("GET", endpoint, data=None, params=params)

    def post_acceptedsenders_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create acceptedSender"""
        endpoint = f"/groups/{group_id}/acceptedSenders/$ref"
        return self.request("POST", endpoint, data=None, params=params)

    def delete_approleassignments_group(
        self, group_id: str, appRoleAssignment_id: str, params: Dict = None
    ) -> Any:
        """Delete appRoleAssignment"""
        endpoint = f"/groups/{group_id}/appRoleAssignments/{appRoleAssignment_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def delta_application(self, params: Dict = None) -> Any:
        """application: delta"""
        endpoint = "/applications/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def list_owneddevices_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List ownedDevices"""
        endpoint = f"/users/{id_or_userPrincipalName}/ownedDevices"
        return self.request("GET", endpoint, data=None, params=params)

    def update_mailboxsettings_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update user mailbox settings"""
        endpoint = f"/users/{id_or_userPrincipalName}/mailboxSettings"
        return self.request("PATCH", endpoint, data=None, params=params)

    def unfollow_driveitem(self, item_id: str, params: Dict = None) -> Any:
        """Unfollow drive item"""
        endpoint = f"/me/drive/following/{item_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def get_user(self, id_or_userPrincipalName: str, params: Dict = None) -> Any:
        """Get a user"""
        endpoint = f"/users/{id_or_userPrincipalName}"
        return self.request("GET", endpoint, data=None, params=params)

    def list_thumbnails_driveitem(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """List thumbnails for a DriveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/thumbnails"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_event(self, event_id: str, params: Dict = None) -> Any:
        """Delete event"""
        endpoint = f"/me/events/{event_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def decline_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: decline"""
        endpoint = f"/me/events/{event_id}/decline"
        return self.request("POST", endpoint, data=None, params=params)

    def list_grouplifecyclepolicies_group(
        self, group_id: str, params: Dict = None
    ) -> Any:
        """List groupLifecyclePolicies"""
        endpoint = f"/groups/{group_id}/groupLifecyclePolicies"
        return self.request("GET", endpoint, data=None, params=params)

    def post_sections_notebook(
        self, notebook_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create section"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}/sections"
        return self.request("POST", endpoint, data=None, params=params)

    def delta_group(self, params: Dict = None) -> Any:
        """group: delta"""
        endpoint = "/groups/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def update_calendar(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update calendar"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendar"
        return self.request("PATCH", endpoint, data=None, params=params)

    def post_threads_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create conversation thread"""
        endpoint = f"/groups/{group_id}/threads"
        return self.request("POST", endpoint, data=None, params=params)

    def checkin_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: checkin"""
        endpoint = f"/drives/{driveId}/items/{itemId}/checkin"
        return self.request("POST", endpoint, data=None, params=params)

    def deleteditems_delete_directory(
        self, deletedItem_id: str, params: Dict = None
    ) -> Any:
        """Permanently delete an item (directory object)"""
        endpoint = f"/directory/deletedItems/{deletedItem_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def changepassword_user(self, data: Dict = None, params: Dict = None) -> Any:
        """user: changePassword"""
        endpoint = "/me/changePassword"
        return self.request("POST", endpoint, data=None, params=params)

    def get_subscription(self, data: Dict = None, params: Dict = None) -> Any:
        """Get subscription"""
        endpoint = "https://graph.microsoft.com/v1.0/subscriptions"
        return self.request("POST", endpoint, data=None, params=params)

    def list_pages_section(self, section_id: str, params: Dict = None) -> Any:
        """List pages"""
        endpoint = f"/me/onenote/sections/{section_id}/pages"
        return self.request("GET", endpoint, data=None, params=params)

    def getbyids_directoryobject(self, data: Dict = None, params: Dict = None) -> Any:
        """directoryObject: getByIds"""
        endpoint = "/directoryObjects/getByIds"
        return self.request("POST", endpoint, data=None, params=params)

    def getschedule_calendar(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """calendar: getSchedule"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendar/getSchedule"
        return self.request("POST", endpoint, data=None, params=params)

    def validateproperties_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """group: validateProperties"""
        endpoint = f"/groups/{group_id}/validateProperties"
        return self.request("POST", endpoint, data=None, params=params)

    def setdata_chart(
        self,
        item_id: str,
        id_or_name: str,
        name: str,
        data: Dict = None,
        params: Dict = None,
    ) -> Any:
        """Chart: setData"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}/setData"
        return self.request("POST", endpoint, data=None, params=params)

    def post_approleassignments_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Grant an appRoleAssignment to a user"""
        endpoint = f"/users/{id_or_userPrincipalName}/appRoleAssignments"
        return self.request("POST", endpoint, data=None, params=params)

    def list_instances_event(self, event_id: str, params: Dict = None) -> Any:
        """List instances"""
        endpoint = f"/me/events/{event_id}/instances"
        return self.request("GET", endpoint, data=None, params=params)

    def overview_resources_calendar(self, params: Dict = None) -> Any:
        """Working with calendars and events using the Microsoft Graph API"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def delete_chart(
        self, item_id: str, id_or_name: str, name: str, params: Dict = None
    ) -> Any:
        """chart: delete"""
        endpoint = (
            f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}"
        )
        return self.request("DELETE", endpoint, data=None, params=params)

    def list_rejectedsenders_group(self, group_id: str, params: Dict = None) -> Any:
        """List rejectedSenders"""
        endpoint = f"/groups/{group_id}/rejectedSenders"
        return self.request("GET", endpoint, data=None, params=params)

    def list_events_user(self, params: Dict = None) -> Any:
        """List events"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def post_groups_group(self, data: Dict = None, params: Dict = None) -> Any:
        """Create group"""
        endpoint = "/groups"
        return self.request("POST", endpoint, data=None, params=params)

    def list_calendargroups_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List calendarGroups"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendarGroups"
        return self.request("GET", endpoint, data=None, params=params)

    def update_subscription(self, data: Dict = None, params: Dict = None) -> Any:
        """Update subscription"""
        endpoint = "https://graph.microsoft.com/v1.0/subscriptions"
        return self.request("POST", endpoint, data=None, params=params)

    def update_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update user"""
        endpoint = f"/users/{id_or_userPrincipalName}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def reminderview_user(
        self,
        id_or_userPrincipalName: str,
        startDateTime_value: str,
        endDateTime_value: str,
        params: Dict = None,
    ) -> Any:
        """user: reminderView"""
        endpoint = f"/users/{id_or_userPrincipalName}/reminderView(startDateTime={startDateTime_value},endDateTime={endDateTime_value})"
        return self.request("GET", endpoint, data=None, params=params)

    def delta_driveitem(self, drive_id: str, params: Dict = None) -> Any:
        """driveItem: delta"""
        endpoint = f"/drives/{drive_id}/root/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def list_acceptedsenders_group(self, group_id: str, params: Dict = None) -> Any:
        """List acceptedSenders"""
        endpoint = f"/groups/{group_id}/acceptedSenders"
        return self.request("GET", endpoint, data=None, params=params)

    def list_approleassignments_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List appRoleAssignments granted to a user"""
        endpoint = f"/users/{id_or_userPrincipalName}/appRoleAssignments"
        return self.request("GET", endpoint, data=None, params=params)

    def deleteditems_getuserownedobjects_directory(
        self, data: Dict = None, params: Dict = None
    ) -> Any:
        """List deleted items (directory objects) owned by a user"""
        endpoint = "/directory/deletedItems/getUserOwnedObjects"
        return self.request("POST", endpoint, data=None, params=params)

    def renew_group(self, group_id: str, data: Dict = None, params: Dict = None) -> Any:
        """group: renew"""
        endpoint = f"/groups/{group_id}/renew"
        return self.request("POST", endpoint, data=None, params=params)

    def permanentdelete_event(
        self, usersId: str, eventId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: permanentDelete"""
        endpoint = f"/users/{usersId}/events/{eventId}/permanentDelete"
        return self.request("POST", endpoint, data=None, params=params)

    def get_event(self, params: Dict = None) -> Any:
        """Get event"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def post_events_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create event"""
        endpoint = f"/groups/{group_id}/events"
        return self.request("POST", endpoint, data=None, params=params)

    def discardcheckout_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: discardCheckout"""
        endpoint = f"/drives/{driveId}/items/{itemId}/discardCheckout"
        return self.request("POST", endpoint, data=None, params=params)

    def delete_user(self, id_or_userPrincipalName: str, params: Dict = None) -> Any:
        """Delete a user"""
        endpoint = f"/users/{id_or_userPrincipalName}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def update_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update group"""
        endpoint = f"/groups/{group_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def delete_permission(
        self, drive_id: str, item_id: str, perm_id: str, params: Dict = None
    ) -> Any:
        """Delete a sharing permission from a file or folder"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/permissions/{perm_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def update_event_group(
        self, group_id: str, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update event"""
        endpoint = f"/groups/{group_id}/events/{event_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def get_group(self, group_id: str, params: Dict = None) -> Any:
        """Get group"""
        endpoint = f"/groups/{group_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def checkmemberobjects_directoryobject(
        self, directoryObject_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """directoryObject: checkMemberObjects"""
        endpoint = f"/directoryObjects/{directoryObject_id}/checkMemberObjects"
        return self.request("POST", endpoint, data=None, params=params)

    def getoffice365groupsactivitystorage_reportroot(
        self, period_value: str, params: Dict = None
    ) -> Any:
        """reportRoot: getOffice365GroupsActivityStorage"""
        endpoint = (
            f"/reports/getOffice365GroupsActivityStorage(period='{period_value}')"
        )
        return self.request("GET", endpoint, data=None, params=params)

    def get_onenotesection(self, section_id: str, params: Dict = None) -> Any:
        """Get section"""
        endpoint = f"/me/onenote/sections/{section_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def list_permissions_driveitem(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """List sharing permissions on a driveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/permissions"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_subscription(self, data: Dict = None, params: Dict = None) -> Any:
        """Delete subscription"""
        endpoint = "https://graph.microsoft.com/v1.0/subscriptions"
        return self.request("POST", endpoint, data=None, params=params)

    def post_events_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create event"""
        endpoint = f"/users/{id_or_userPrincipalName}/events"
        return self.request("POST", endpoint, data=None, params=params)

    def list_directreports_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List directReports"""
        endpoint = f"/users/{id_or_userPrincipalName}/directReports"
        return self.request("GET", endpoint, data=None, params=params)

    def exportpersonaldata_user(
        self, user_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """user: exportPersonalData"""
        endpoint = f"/users/{user_id}/exportPersonalData"
        return self.request("POST", endpoint, data=None, params=params)

    def list_sectiongroups_notebook(self, notebook_id: str, params: Dict = None) -> Any:
        """List sectionGroups"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
        return self.request("GET", endpoint, data=None, params=params)

    def checkout_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: checkout"""
        endpoint = f"/drives/{driveId}/items/{itemId}/checkout"
        return self.request("POST", endpoint, data=None, params=params)

    def list_ownedobjects_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List ownedObjects"""
        endpoint = f"/users/{id_or_userPrincipalName}/ownedObjects"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_acceptedsenders_group(self, group_id: str, params: Dict = None) -> Any:
        """Remove acceptedSender"""
        endpoint = f"/groups/{group_id}/acceptedSenders/$ref"
        return self.request("DELETE", endpoint, data=None, params=params)

    def get_event_group(self, group_id: str, event_id: str, params: Dict = None) -> Any:
        """Get event"""
        endpoint = f"/groups/{group_id}/events/{event_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def compute_userprotectionscopecontainer(
        self, data: Dict = None, params: Dict = None
    ) -> Any:
        """userProtectionScopeContainer: compute"""
        endpoint = "/me/dataSecurityAndGovernance/protectionScopes/compute"
        return self.request("POST", endpoint, data=None, params=params)

    def setposition_chart(
        self,
        item_id: str,
        id_or_name: str,
        name: str,
        data: Dict = None,
        params: Dict = None,
    ) -> Any:
        """Chart: setPosition"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}/setPosition"
        return self.request("POST", endpoint, data=None, params=params)

    def list_createdobjects_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List createdObjects"""
        endpoint = f"/users/{id_or_userPrincipalName}/createdObjects"
        return self.request("GET", endpoint, data=None, params=params)

    def delta_directoryobject(self, params: Dict = None) -> Any:
        """directoryObject: delta"""
        endpoint = "/directoryObjects/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def post_subscriptions_subscription(
        self, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create subscription"""
        endpoint = "https://graph.microsoft.com/v1.0/subscriptions"
        return self.request("POST", endpoint, data=None, params=params)

    def deleteditems_get_directory(self, object_id: str, params: Dict = None) -> Any:
        """Get deleted item (directory object)"""
        endpoint = f"/directory/deletedItems/{object_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def list_calendarview_group(self, group_id: str, params: Dict = None) -> Any:
        """List group calendarView"""
        endpoint = f"/groups/{group_id}/calendarView"
        return self.request("GET", endpoint, data=None, params=params)

    def search_driveitem(
        self, drive_id: str, search_text: str, params: Dict = None
    ) -> Any:
        """Search for DriveItems within a drive"""
        endpoint = f"/drives/{drive_id}/root/search(q='{search_text}')"
        return self.request("GET", endpoint, data=None, params=params)

    def image_chart(
        self, item_id: str, id_or_name: str, name: str, params: Dict = None
    ) -> Any:
        """Chart: Image"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}/image"
        return self.request("GET", endpoint, data=None, params=params)

    def overview_resources_groups(self, params: Dict = None) -> Any:
        """Manage groups in Microsoft Graph"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def post_users_user(self, data: Dict = None, params: Dict = None) -> Any:
        """Create User"""
        endpoint = "/users"
        return self.request("POST", endpoint, data=None, params=params)

    def post_sectiongroups_notebook(
        self, notebook_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create sectionGroup"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
        return self.request("POST", endpoint, data=None, params=params)

    def delta_user(self, params: Dict = None) -> Any:
        """user: delta"""
        endpoint = "/users/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def createlink_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create a sharing link for a DriveItem"""
        endpoint = f"/drives/{driveId}/items/{itemId}/createLink"
        return self.request("POST", endpoint, data=None, params=params)

    def delta_serviceprincipal(self, params: Dict = None) -> Any:
        """servicePrincipal: delta"""
        endpoint = "/servicePrincipals/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def get_itemanalytics(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """Get itemAnalytics"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/analytics/allTime"
        return self.request("GET", endpoint, data=None, params=params)

    def getoffice365groupsactivitydetail_reportroot(
        self, period_value: str, params: Dict = None
    ) -> Any:
        """reportRoot: getOffice365GroupsActivityDetail"""
        endpoint = f"/reports/getOffice365GroupsActivityDetail(period='{period_value}')"
        return self.request("GET", endpoint, data=None, params=params)

    def edge_api_overview_resources_browser(self, params: Dict = None) -> Any:
        """Use the Edge API in Microsoft Graph"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def notifications_api_overview_resources_change(self, params: Dict = None) -> Any:
        """Microsoft Graph API change notifications"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def list_calendarview_user(self, params: Dict = None) -> Any:
        """List calendarView"""
        endpoint = "/me/calendar/calendarView"
        return self.request("GET", endpoint, data=None, params=params)

    def checkmembergroups_directoryobject(
        self, directoryObject_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """directoryObject: checkMemberGroups"""
        endpoint = f"/directoryObjects/{directoryObject_id}/checkMemberGroups"
        return self.request("POST", endpoint, data=None, params=params)

    def post_calendargroups_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create CalendarGroup"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendarGroups"
        return self.request("POST", endpoint, data=None, params=params)

    def reauthorize_subscription(self, data: Dict = None, params: Dict = None) -> Any:
        """subscription: reauthorize"""
        endpoint = "https://graph.microsoft.com/v1.0/subscriptions"
        return self.request("POST", endpoint, data=None, params=params)

    def list_calendars_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List calendars"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendars"
        return self.request("GET", endpoint, data=None, params=params)

    def put_content_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Upload or replace the contents of a driveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/content"
        return self.request("PUT", endpoint, data=None, params=params)

    def list_versions_driveitem(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """List versions"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/versions"
        return self.request("GET", endpoint, data=None, params=params)

    def get_driveitem(self, drive_id: str, item_id: str, params: Dict = None) -> Any:
        """Get driveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def update_thread_group(
        self, group_id: str, thread_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update conversation thread"""
        endpoint = f"/groups/{group_id}/threads/{thread_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def list_chart(self, item_id: str, id_or_name: str, params: Dict = None) -> Any:
        """List ChartCollection"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts"
        return self.request("GET", endpoint, data=None, params=params)

    def post_members_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Add members"""
        endpoint = f"/groups/{group_id}/members/$ref"
        return self.request("POST", endpoint, data=None, params=params)

    def processcontent_userdatasecurityandgovernance(
        self, data: Dict = None, params: Dict = None
    ) -> Any:
        """userDataSecurityAndGovernance: processContent"""
        endpoint = "/me/dataSecurityAndGovernance/processContent"
        return self.request("POST", endpoint, data=None, params=params)

    def post_series_chart(
        self,
        item_id: str,
        id_or_name: str,
        name: str,
        data: Dict = None,
        params: Dict = None,
    ) -> Any:
        """Create ChartSeries"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}/series"
        return self.request("POST", endpoint, data=None, params=params)

    def socketio_subscriptions(self, driveId: str, params: Dict = None) -> Any:
        """Get websocket endpoint"""
        endpoint = f"/drives/{driveId}/root/subscriptions/socketIo"
        return self.request("GET", endpoint, data=None, params=params)

    def getoffice365groupsactivitycounts_reportroot(
        self, period_value: str, params: Dict = None
    ) -> Any:
        """reportRoot: getOffice365GroupsActivityCounts"""
        endpoint = f"/reports/getOffice365GroupsActivityCounts(period='{period_value}')"
        return self.request("GET", endpoint, data=None, params=params)

    def list_tasks_planneruser(self, user_id: str, params: Dict = None) -> Any:
        """List tasks"""
        endpoint = f"/users/{user_id}/planner/tasks"
        return self.request("GET", endpoint, data=None, params=params)

    def findmeetingtimes_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """user: findMeetingTimes"""
        endpoint = f"/users/{id_or_userPrincipalName}/findMeetingTimes"
        return self.request("POST", endpoint, data=None, params=params)

    def list_oauth2permissiongrants_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List a user's oauth2PermissionGrants"""
        endpoint = f"/users/{id_or_userPrincipalName}/oauth2PermissionGrants"
        return self.request("GET", endpoint, data=None, params=params)

    def list_people_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List people"""
        endpoint = f"/users/{id_or_userPrincipalName}/people"
        return self.request("GET", endpoint, data=None, params=params)

    def post_conversations_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create conversation"""
        endpoint = f"/groups/{group_id}/conversations"
        return self.request("POST", endpoint, data=None, params=params)

    def delta_event(self, params: Dict = None) -> Any:
        """event: delta"""
        endpoint = "/me/calendarView/delta"
        return self.request("GET", endpoint, data=None, params=params)

    def list_approleassignments_group(self, group_id: str, params: Dict = None) -> Any:
        """List appRoleAssignments granted to a group"""
        endpoint = f"/groups/{group_id}/appRoleAssignments"
        return self.request("GET", endpoint, data=None, params=params)

    def cancel_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: cancel"""
        endpoint = f"/me/events/{event_id}/cancel"
        return self.request("POST", endpoint, data=None, params=params)

    def list_children_driveitem(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """List children of a driveItem"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        return self.request("GET", endpoint, data=None, params=params)

    def post_pages_section(
        self, section_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create page"""
        endpoint = f"/me/onenote/sections/{section_id}/pages"
        return self.request("POST", endpoint, data=None, params=params)

    def list_plans_plannergroup(self, group_id: str, params: Dict = None) -> Any:
        """List plans"""
        endpoint = f"/groups/{group_id}/planner/plans"
        return self.request("GET", endpoint, data=None, params=params)

    def post_owners_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Add owners"""
        endpoint = f"/groups/{group_id}/owners/$ref"
        return self.request("POST", endpoint, data=None, params=params)

    def delete_approleassignments_user(
        self, user_id: str, appRoleAssignment_id: str, params: Dict = None
    ) -> Any:
        """Delete appRoleAssignment"""
        endpoint = f"/users/{user_id}/appRoleAssignments/{appRoleAssignment_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def list_registereddevices_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List registeredDevices"""
        endpoint = f"/users/{id_or_userPrincipalName}/registeredDevices"
        return self.request("GET", endpoint, data=None, params=params)

    def upsert_group(self, data: Dict = None, params: Dict = None) -> Any:
        """Upsert group"""
        endpoint = "/groups(uniqueName='uniqueName')"
        return self.request("PATCH", endpoint, data=None, params=params)

    def list_group(self, params: Dict = None) -> Any:
        """List groups"""
        endpoint = "/groups"
        return self.request("GET", endpoint, data=None, params=params)

    def update_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Update event"""
        endpoint = f"/me/events/{event_id}"
        return self.request("PATCH", endpoint, data=None, params=params)

    def delete_group(self, group_id: str, params: Dict = None) -> Any:
        """Delete group"""
        endpoint = f"/groups/{group_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def getmembergroups_directoryobject(
        self, directoryObject_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """directoryObject: getMemberGroups"""
        endpoint = f"/directoryObjects/{directoryObject_id}/getMemberGroups"
        return self.request("POST", endpoint, data=None, params=params)

    def list_events_group(self, group_id: str, params: Dict = None) -> Any:
        """List events"""
        endpoint = f"/groups/{group_id}/events"
        return self.request("GET", endpoint, data=None, params=params)

    def getnotebookfromweburl_notebook(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """notebook: getNotebookFromWebUrl"""
        endpoint = (
            f"/users/{id_or_userPrincipalName}/onenote/notebooks/GetNotebookFromWebUrl"
        )
        return self.request("POST", endpoint, data=None, params=params)

    def callrecord_resources_callrecords(self, params: Dict = None) -> Any:
        """callRecord resource type"""
        endpoint = ""
        return self.request("GET", endpoint, data=None, params=params)

    def deleteditems_restore_directory(
        self, deletedItem_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Restore deleted item (directory object)"""
        endpoint = f"/directory/deletedItems/{deletedItem_id}/restore"
        return self.request("POST", endpoint, data=None, params=params)

    def get_content_driveitem(
        self, drive_id: str, item_id: str, params: Dict = None
    ) -> Any:
        """Download driveItem content"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/content"
        return self.request("GET", endpoint, data=None, params=params)

    def post_calendars_user(
        self, id_or_userPrincipalName: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create calendar"""
        endpoint = f"/users/{id_or_userPrincipalName}/calendars"
        return self.request("POST", endpoint, data=None, params=params)

    def get_opentypeextension(
        self, Id: str, extensionId: str, params: Dict = None
    ) -> Any:
        """Get open extension"""
        endpoint = f"/devices/{Id}/extensions/{extensionId}"
        return self.request("GET", endpoint, data=None, params=params)

    def itemat_chartcollection(
        self, item_id: str, id_or_name: str, index: str, params: Dict = None
    ) -> Any:
        """ChartCollection: ItemAt"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/itemAt(index={index})"
        return self.request("GET", endpoint, data=None, params=params)

    def invite_driveitem(
        self, drive_id: str, item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: invite"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/invite"
        return self.request("POST", endpoint, data=None, params=params)

    def getrecentnotebooks_notebook(
        self, includePersonalNotebooks: str, params: Dict = None
    ) -> Any:
        """notebook: getRecentNotebooks"""
        endpoint = f"/me/onenote/notebooks/getRecentNotebooks(includePersonalNotebooks={includePersonalNotebooks})"
        return self.request("GET", endpoint, data=None, params=params)

    def list_transitivemembers_group(self, group_id: str, params: Dict = None) -> Any:
        """List group transitive members"""
        endpoint = f"/groups/{group_id}/transitiveMembers"
        return self.request("GET", endpoint, data=None, params=params)

    def deleteditems_list_directory(self, params: Dict = None) -> Any:
        """List deletedItems (directory objects)"""
        endpoint = "/directory/deletedItems/microsoft.graph.administrativeUnit"
        return self.request("GET", endpoint, data=None, params=params)

    def get_notebook(self, notebook_id: str, params: Dict = None) -> Any:
        """Get notebook"""
        endpoint = f"/me/onenote/notebooks/{notebook_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def get_content_format_driveitem(self, item_id: str, params: Dict = None) -> Any:
        """Download a file in another format"""
        endpoint = f"/drive/items/{item_id}/content"
        return self.request("GET", endpoint, data=None, params=params)

    def post_contentactivities_activitiescontainer(
        self, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create contentActivity"""
        endpoint = "/me/dataSecurityAndGovernance/activities/contentActivities"
        return self.request("POST", endpoint, data=None, params=params)

    def get_chart(
        self, item_id: str, id_or_name: str, name: str, params: Dict = None
    ) -> Any:
        """Get Chart"""
        endpoint = (
            f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}"
        )
        return self.request("GET", endpoint, data=None, params=params)

    def tentativelyaccept_event(
        self, event_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """event: tentativelyAccept"""
        endpoint = f"/me/events/{event_id}/tentativelyAccept"
        return self.request("POST", endpoint, data=None, params=params)

    def post_children_driveitem(
        self, drive_id: str, parent_item_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Create a new folder in a drive"""
        endpoint = f"/drives/{drive_id}/items/{parent_item_id}/children"
        return self.request("POST", endpoint, data=None, params=params)

    def copy_driveitem(
        self, driveId: str, itemId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """driveItem: copy"""
        endpoint = f"/drives/{driveId}/items/{itemId}/copy"
        return self.request("POST", endpoint, data=None, params=params)

    def copytosectiongroup_section(
        self, section_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """section: copyToSectionGroup"""
        endpoint = f"/me/onenote/sections/{section_id}/copyToSectionGroup"
        return self.request("POST", endpoint, data=None, params=params)

    def get_drive(self, params: Dict = None) -> Any:
        """Get Drive"""
        endpoint = "/me/drive"
        return self.request("GET", endpoint, data=None, params=params)

    def list_series_chart(
        self, item_id: str, id_or_name: str, name: str, params: Dict = None
    ) -> Any:
        """List series"""
        endpoint = f"/me/drive/items/{item_id}/workbook/worksheets/{id_or_name}/charts/{name}/series"
        return self.request("GET", endpoint, data=None, params=params)

    def post_approleassignments_group(
        self, groupId: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """Grant an appRoleAssignment to a group"""
        endpoint = f"/groups/{groupId}/appRoleAssignments"
        return self.request("POST", endpoint, data=None, params=params)

    def delete_thread_group(
        self, group_id: str, thread_id: str, params: Dict = None
    ) -> Any:
        """Delete conversation thread"""
        endpoint = f"/groups/{group_id}/threads/{thread_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def delete_conversation_group(
        self, group_id: str, conversation_id: str, params: Dict = None
    ) -> Any:
        """Delete conversation"""
        endpoint = f"/groups/{group_id}/conversations/{conversation_id}"
        return self.request("DELETE", endpoint, data=None, params=params)

    def reply_conversationthread(
        self, group_id: str, thread_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """conversationThread: reply"""
        endpoint = f"/groups/{group_id}/threads/{thread_id}/reply"
        return self.request("POST", endpoint, data=None, params=params)

    def getactivitybyinterval_itemactivitystat(
        self,
        drive_id: str,
        item_id: str,
        startDateTime: str,
        endDateTime: str,
        interval: str,
        params: Dict = None,
    ) -> Any:
        """Get item activity stats by interval"""
        endpoint = f"/drives/{drive_id}/items/{item_id}/getActivitiesByInterval(startDateTime={startDateTime},endDateTime={endDateTime},interval={interval})"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_owners_group(
        self, group_id: str, owner_id: str, params: Dict = None
    ) -> Any:
        """Remove group owner"""
        endpoint = f"/groups/{group_id}/owners/{owner_id}/$ref"
        return self.request("DELETE", endpoint, data=None, params=params)

    def assignlicense_group(
        self, group_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """group: assignLicense"""
        endpoint = f"/groups/{group_id}/assignLicense"
        return self.request("POST", endpoint, data=None, params=params)

    def getoffice365groupsactivityfilecounts_reportroot(
        self, period_value: str, params: Dict = None
    ) -> Any:
        """reportRoot: getOffice365GroupsActivityFileCounts"""
        endpoint = (
            f"/reports/getOffice365GroupsActivityFileCounts(period='{period_value}')"
        )
        return self.request("GET", endpoint, data=None, params=params)

    def get_thread_group(
        self, group_id: str, thread_id: str, params: Dict = None
    ) -> Any:
        """Get conversation thread"""
        endpoint = f"/groups/{group_id}/threads/{thread_id}"
        return self.request("GET", endpoint, data=None, params=params)

    def retryserviceprovisioning_user(
        self, user_id: str, data: Dict = None, params: Dict = None
    ) -> Any:
        """user: retryServiceProvisioning"""
        endpoint = f"/users/{user_id}/retryServiceProvisioning"
        return self.request("POST", endpoint, data=None, params=params)

    def list_manager_user(
        self, id_or_userPrincipalName: str, params: Dict = None
    ) -> Any:
        """List manager"""
        endpoint = f"/users/{id_or_userPrincipalName}/manager"
        return self.request("GET", endpoint, data=None, params=params)

    def delete_event_group(
        self, group_id: str, event_id: str, params: Dict = None
    ) -> Any:
        """Delete event"""
        endpoint = f"/groups/{group_id}/events/{event_id}"
        return self.request("DELETE", endpoint, data=None, params=params)
