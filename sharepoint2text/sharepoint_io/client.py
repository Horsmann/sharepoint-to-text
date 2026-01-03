"""
SharePoint client using Microsoft Graph API with Entra ID app authentication.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from typing import Any, Callable, Iterator
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, urlparse
from urllib.request import Request, urlopen

from sharepoint2text.sharepoint_io.exceptions import (
    SharePointAuthError,
    SharePointRequestError,
)

_TOKEN_ENDPOINT_TEMPLATE = (
    "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
)
_GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class EntraIDAppCredentials:
    """Client credentials for an Entra ID application."""

    tenant_id: str
    client_id: str
    client_secret: str
    scope: str = "https://graph.microsoft.com/.default"


# System fields returned by Graph API that are not custom columns
_SYSTEM_FIELDS = frozenset(
    {
        "@odata.etag",
        "id",
        "ContentType",
        "Created",
        "Modified",
        "AuthorLookupId",
        "EditorLookupId",
        "_UIVersionString",
        "Attachments",
        "Edit",
        "LinkFilenameNoMenu",
        "LinkFilename",
        "DocIcon",
        "ItemChildCount",
        "FolderChildCount",
        "_ComplianceFlags",
        "_ComplianceTag",
        "_ComplianceTagWrittenTime",
        "_ComplianceTagUserId",
        "_CommentCount",
        "_LikeCount",
        "_DisplayName",
        "FileLeafRef",
        "FileDirRef",
        "FileRef",
        "_CheckinComment",
        "LinkTitleNoMenu",
        "LinkTitle",
        "_IsRecord",
        "_VirusStatus",
        "_VirusVendorID",
        "_VirusInfo",
        "SharedWithUsersId",
        "SharedWithDetails",
        "Restricted",
        # Additional system/derived fields
        "FileSizeDisplay",
        "ParentVersionStringLookupId",
        "ParentLeafNameLookupId",
        "Title",
        "_ExtendedDescription",
        "CheckoutUserLookupId",
        "CheckedOutUserId",
        "IsCheckedoutToLocal",
        "_CopySource",
        "_HasCopyDestinations",
        "TemplateUrl",
        "xd_ProgID",
        "xd_Signature",
        "Order",
        "GUID",
        "WorkflowVersion",
        "WorkflowInstanceID",
        "AccessPolicy",
        "BSN",
        "HTML_x0020_File_x0020_Type",
        "_SourceUrl",
        "_SharedFileIndex",
        "MetaInfo",
        "_Level",
        "ProgId",
        "ScopeId",
        "A2ODMountCount",
        "SyncClientId",
        "_ShortcutUrl",
        "_ShortcutSiteId",
        "_ShortcutWebId",
        "_ShortcutUniqueId",
    }
)


@dataclass(frozen=True)
class SharePointFileMetadata:
    """Metadata for a file stored in SharePoint."""

    name: str
    id: str
    web_url: str
    download_url: str | None = None
    size: int | None = None
    mime_type: str | None = None
    last_modified: str | None = None
    created: str | None = None
    parent_path: str | None = None
    custom_fields: dict[str, Any] | None = None


class SharePointRestClient:
    """SharePoint client using Microsoft Graph API."""

    def __init__(
        self,
        site_url: str,
        credentials: EntraIDAppCredentials,
        *,
        request_func: Callable[..., object] | None = None,
        timeout: float = 30.0,
    ) -> None:
        self._site_url = site_url.rstrip("/")
        self._credentials = credentials
        self._request = request_func or urlopen
        self._timeout = timeout
        self._access_token: str | None = None
        self._site_id: str | None = None

    def fetch_access_token(self) -> str:
        """Request an app-only access token from Entra ID."""
        token_url = _TOKEN_ENDPOINT_TEMPLATE.format(
            tenant_id=self._credentials.tenant_id
        )
        logger.info(f"Fetching access token from {token_url}")
        payload = urlencode(
            {
                "client_id": self._credentials.client_id,
                "client_secret": self._credentials.client_secret,
                "scope": self._credentials.scope,
                "grant_type": "client_credentials",
            }
        ).encode("utf-8")
        request = Request(
            token_url,
            data=payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        _, body = self._send(request, request_kind="token")
        try:
            data = json.loads(body.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise SharePointAuthError("Invalid token response JSON") from exc
        access_token = data.get("access_token")
        if not access_token:
            raise SharePointAuthError("Token response missing access_token")
        self._access_token = access_token
        return access_token

    def _ensure_token(self) -> str:
        """Ensure we have a valid access token."""
        if self._access_token is None:
            self.fetch_access_token()
        return self._access_token  # type: ignore[return-value]

    def _get_headers(self) -> dict[str, str]:
        """Get request headers with authorization."""
        token = self._ensure_token()
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

    def get_site_id(self) -> str:
        """Get the Graph API site ID from the site URL."""
        if self._site_id is not None:
            return self._site_id

        parsed = urlparse(self._site_url)
        hostname = parsed.netloc
        site_path = parsed.path.rstrip("/")

        # Graph API endpoint to get site by hostname and path
        # Format: /sites/{hostname}:/{site-path}
        if site_path:
            url = f"{_GRAPH_API_BASE}/sites/{hostname}:{site_path}"
        else:
            url = f"{_GRAPH_API_BASE}/sites/{hostname}"

        data = self._get_json(url)
        site_id = data.get("id")
        if not isinstance(site_id, str):
            raise SharePointRequestError(
                "Could not get site ID from Graph API",
                status_code=None,
                body=json.dumps(data),
                url=url,
            )
        self._site_id = site_id
        logger.info(f"Resolved site ID: {site_id}")
        return site_id

    def list_all_files(
        self,
        *,
        include_root_files: bool = True,
    ) -> list[SharePointFileMetadata]:
        """
        List all files in the SharePoint site's default document library.

        Uses Microsoft Graph API to traverse the document library.
        """
        site_id = self.get_site_id()
        files: list[SharePointFileMetadata] = []

        # List all files recursively from the root
        for file_meta in self._walk_drive_items(site_id, item_id=None):
            files.append(file_meta)

        return files

    def list_drives(self) -> list[dict[str, Any]]:
        """List all document libraries (drives) in the site."""
        site_id = self.get_site_id()
        url = f"{_GRAPH_API_BASE}/sites/{site_id}/drives"
        data = self._get_json(url)
        return data.get("value", [])

    def list_files_in_folder(
        self,
        folder_path: str = "/",
        *,
        drive_id: str | None = None,
    ) -> list[SharePointFileMetadata]:
        """List files in a specific folder."""
        site_id = self.get_site_id()

        if drive_id is None:
            # Use default drive
            if folder_path == "/" or not folder_path:
                base = f"{_GRAPH_API_BASE}/sites/{site_id}/drive/root/children"
            else:
                encoded_path = quote(folder_path.strip("/"), safe="")
                base = f"{_GRAPH_API_BASE}/sites/{site_id}/drive/root:/{encoded_path}:/children"
        else:
            if folder_path == "/" or not folder_path:
                base = (
                    f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root/children"
                )
            else:
                encoded_path = quote(folder_path.strip("/"), safe="")
                base = f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/children"

        url = f"{base}?$expand=listItem($expand=fields)"
        return list(self._list_items_paginated(url))

    def download_file(self, file_id: str, *, drive_id: str | None = None) -> bytes:
        """Download a file by its ID and return its bytes."""
        site_id = self.get_site_id()

        if drive_id is None:
            url = f"{_GRAPH_API_BASE}/sites/{site_id}/drive/items/{file_id}/content"
        else:
            url = f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"

        request = Request(
            url,
            headers=self._get_headers(),
            method="GET",
        )
        _, body = self._send(request, request_kind="file download")
        return body

    def download_file_by_path(
        self, file_path: str, *, drive_id: str | None = None
    ) -> bytes:
        """Download a file by its path and return its bytes."""
        site_id = self.get_site_id()
        encoded_path = quote(file_path.strip("/"), safe="/")

        if drive_id is None:
            url = (
                f"{_GRAPH_API_BASE}/sites/{site_id}/drive/root:/{encoded_path}:/content"
            )
        else:
            url = f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}:/content"

        request = Request(
            url,
            headers=self._get_headers(),
            method="GET",
        )
        _, body = self._send(request, request_kind="file download")
        return body

    def _build_children_url(
        self,
        site_id: str,
        item_id: str | None,
        drive_id: str | None = None,
    ) -> str:
        """Build URL for listing children of a drive item."""
        if drive_id is None:
            if item_id is None:
                base = f"{_GRAPH_API_BASE}/sites/{site_id}/drive/root/children"
            else:
                base = (
                    f"{_GRAPH_API_BASE}/sites/{site_id}/drive/items/{item_id}/children"
                )
        else:
            if item_id is None:
                base = (
                    f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root/children"
                )
            else:
                base = f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/items/{item_id}/children"
        # Expand listItem.fields to get custom column values
        return f"{base}?$expand=listItem($expand=fields)"

    def _walk_drive_items(
        self,
        site_id: str,
        item_id: str | None,
        *,
        drive_id: str | None = None,
        parent_path: str = "",
    ) -> Iterator[SharePointFileMetadata]:
        """Recursively walk through drive items and yield file metadata."""
        url = self._build_children_url(site_id, item_id, drive_id)

        for item in self._list_items_paginated(url, parent_path=parent_path):
            yield item

        # We need to get the folders separately to recurse into them
        for item in self._get_folders_from_url(url):
            folder_name = item.get("name", "")
            folder_id = item.get("id")
            new_parent_path = (
                f"{parent_path}/{folder_name}" if parent_path else folder_name
            )
            if folder_id:
                yield from self._walk_drive_items(
                    site_id,
                    folder_id,
                    drive_id=drive_id,
                    parent_path=new_parent_path,
                )

    def _get_folders_from_url(self, url: str) -> list[dict[str, Any]]:
        """Get folder items from a URL."""
        folders = []
        current_url: str | None = url

        while current_url:
            data = self._get_json(current_url)
            items = data.get("value", [])

            for item in items:
                if isinstance(item, dict) and "folder" in item:
                    folders.append(item)

            current_url = data.get("@odata.nextLink")

        return folders

    def _list_items_paginated(
        self, url: str, *, parent_path: str = ""
    ) -> Iterator[SharePointFileMetadata]:
        """List items with pagination support, yielding only files."""
        current_url: str | None = url

        while current_url:
            data = self._get_json(current_url)
            items = data.get("value", [])

            for item in items:
                if not isinstance(item, dict):
                    continue

                # Skip folders, only yield files
                if "folder" in item:
                    continue

                # This is a file
                if "file" in item:
                    yield self._parse_file_item(item, parent_path)

            # Handle pagination
            current_url = data.get("@odata.nextLink")

    def _parse_file_item(
        self, item: dict[str, Any], parent_path: str = ""
    ) -> SharePointFileMetadata:
        """Parse a Graph API drive item into SharePointFileMetadata."""
        file_info = item.get("file", {})

        # Extract custom fields from listItem.fields
        custom_fields = self._extract_custom_fields(item)

        return SharePointFileMetadata(
            name=item.get("name", ""),
            id=item.get("id", ""),
            web_url=item.get("webUrl", ""),
            download_url=item.get("@microsoft.graph.downloadUrl"),
            size=item.get("size"),
            mime_type=(
                file_info.get("mimeType") if isinstance(file_info, dict) else None
            ),
            last_modified=item.get("lastModifiedDateTime"),
            created=item.get("createdDateTime"),
            parent_path=parent_path or None,
            custom_fields=custom_fields if custom_fields else None,
        )

    def _extract_custom_fields(self, item: dict[str, Any]) -> dict[str, Any]:
        """Extract custom column values from listItem.fields."""
        list_item = item.get("listItem")
        if not isinstance(list_item, dict):
            return {}

        fields = list_item.get("fields")
        if not isinstance(fields, dict):
            return {}

        # Filter out system fields to get only custom columns
        custom = {}
        for key, value in fields.items():
            if key not in _SYSTEM_FIELDS and not key.startswith("@odata"):
                custom[key] = value

        return custom

    def _get_json(self, url: str) -> dict[str, Any]:
        """Make a GET request and parse JSON response."""
        request = Request(url, headers=self._get_headers(), method="GET")
        _, body = self._send(request, request_kind="API")
        text = body.decode("utf-8", errors="replace")
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise SharePointRequestError(
                "Invalid JSON response from Graph API",
                status_code=None,
                body=text,
                url=url,
            ) from exc

    def _send(self, request: Request, *, request_kind: str) -> tuple[int, bytes]:
        """Send an HTTP request and return status code and body."""
        response = None
        try:
            response = self._request(request, timeout=self._timeout)
        except HTTPError as exc:
            body = exc.read()
            raise SharePointRequestError(
                f"{request_kind} request failed with status {exc.code}",
                status_code=exc.code,
                body=body.decode("utf-8", errors="replace") if body else None,
                url=request.full_url,
            ) from exc
        except URLError as exc:
            raise SharePointRequestError(
                f"{request_kind} request failed due to network error: {exc.reason}",
                status_code=None,
                body=None,
                url=request.full_url,
            ) from exc

        try:
            status = getattr(response, "status", None)
            if status is None:
                status = response.getcode()
            body = response.read()
        finally:
            if response is not None:
                try:
                    response.close()
                except Exception:
                    pass

        if status is None or not (200 <= status < 300):
            raise SharePointRequestError(
                f"{request_kind} request returned status {status}",
                status_code=status,
                body=body.decode("utf-8", errors="replace") if body else None,
                url=request.full_url,
            )
        return status, body
