"""
SharePoint client using Microsoft Graph API with Entra ID app authentication.
"""

from __future__ import annotations

import fnmatch
import json
import logging
from dataclasses import dataclass, field
from datetime import datetime
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

    def get_full_path(self) -> str:
        """Get the full path of the file including parent path."""
        if self.parent_path:
            return f"{self.parent_path}/{self.name}"
        return self.name


@dataclass
class FileFilter:
    """
    Filter criteria for selecting files from SharePoint.

    Use this to implement delta-sync patterns by filtering on dates,
    or to target specific folders/paths within a SharePoint site.

    Attributes:
        created_after: Only include files created after this datetime (inclusive)
        created_before: Only include files created before this datetime (exclusive)
        modified_after: Only include files modified after this datetime (inclusive)
        modified_before: Only include files modified before this datetime (exclusive)
        folder_paths: List of specific folder paths to search (e.g., ["Documents/Reports"])
            If empty, searches entire drive. Paths are relative to drive root.
        path_patterns: Glob-style patterns to match against full file paths
            (e.g., ["*.docx", "Reports/**/*.pdf", "2024-*/*"])
            Patterns are matched against the full path including parent folders.
        extensions: List of file extensions to include (e.g., [".docx", ".pdf"])
            Extensions should include the leading dot.

    Example:
        # Get all PDFs modified in the last 7 days from Reports folder
        filter = FileFilter(
            modified_after=datetime.now(timezone.utc) - timedelta(days=7),
            folder_paths=["Documents/Reports"],
            extensions=[".pdf"],
        )

        # Get all files matching a pattern
        filter = FileFilter(
            path_patterns=["Projects/2024-*/**/*.xlsx"],
        )
    """

    created_after: datetime | None = None
    created_before: datetime | None = None
    modified_after: datetime | None = None
    modified_before: datetime | None = None
    folder_paths: list[str] = field(default_factory=list)
    path_patterns: list[str] = field(default_factory=list)
    extensions: list[str] = field(default_factory=list)

    def matches(self, file_meta: SharePointFileMetadata) -> bool:
        """
        Check if a file matches all filter criteria.

        Args:
            file_meta: The file metadata to check

        Returns:
            True if the file matches all specified criteria, False otherwise
        """
        # Check creation date filters
        if self.created_after or self.created_before:
            if not file_meta.created:
                return False
            created_dt = _parse_iso_datetime(file_meta.created)
            if created_dt is None:
                return False
            if self.created_after and created_dt < self.created_after:
                return False
            if self.created_before and created_dt >= self.created_before:
                return False

        # Check modification date filters
        if self.modified_after or self.modified_before:
            if not file_meta.last_modified:
                return False
            modified_dt = _parse_iso_datetime(file_meta.last_modified)
            if modified_dt is None:
                return False
            if self.modified_after and modified_dt < self.modified_after:
                return False
            if self.modified_before and modified_dt >= self.modified_before:
                return False

        # Check extension filter
        if self.extensions:
            name_lower = file_meta.name.lower()
            if not any(name_lower.endswith(ext.lower()) for ext in self.extensions):
                return False

        # Check path pattern filters
        if self.path_patterns:
            full_path = file_meta.get_full_path()
            if not any(
                fnmatch.fnmatch(full_path, pattern) for pattern in self.path_patterns
            ):
                return False

        return True

    def get_target_folders(self) -> list[str]:
        """
        Get the list of folders to search.

        Returns:
            List of folder paths to search, or empty list to search entire drive
        """
        return self.folder_paths


def _parse_iso_datetime(dt_string: str) -> datetime | None:
    """
    Parse an ISO 8601 datetime string to a datetime object.

    Handles formats like: 2024-01-15T10:30:00Z, 2024-01-15T10:30:00.123Z
    """
    try:
        # Handle the 'Z' suffix (UTC)
        if dt_string.endswith("Z"):
            dt_string = dt_string[:-1] + "+00:00"
        # Python 3.11+ has fromisoformat that handles this, but for compatibility:
        # Remove microseconds if present for simpler parsing
        if "." in dt_string:
            # Split at the dot and reconstruct with timezone
            base, rest = dt_string.split(".", 1)
            # Find timezone info (+ or - after the dot)
            tz_idx = rest.find("+")
            if tz_idx == -1:
                tz_idx = rest.find("-")
            if tz_idx != -1:
                dt_string = base + rest[tz_idx:]
            else:
                dt_string = base
        return datetime.fromisoformat(dt_string)
    except (ValueError, AttributeError):
        return None


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

    def list_files_filtered(
        self,
        file_filter: FileFilter,
        *,
        drive_id: str | None = None,
    ) -> Iterator[SharePointFileMetadata]:
        """
        List files matching the specified filter criteria.

        This method enables delta-sync patterns by allowing filtering on creation
        and modification dates, as well as targeting specific folders or path patterns.

        When folder_paths are specified in the filter, only those folders are searched,
        which can significantly reduce API calls for large document libraries.

        Args:
            file_filter: Filter criteria to apply
            drive_id: Optional drive ID to search. If None, uses the default drive.

        Yields:
            SharePointFileMetadata for each file matching the filter criteria

        Example:
            # Delta sync: get files modified in the last 24 hours
            from datetime import datetime, timedelta, timezone

            filter = FileFilter(
                modified_after=datetime.now(timezone.utc) - timedelta(hours=24),
            )
            for file in client.list_files_filtered(filter):
                print(f"Modified: {file.name}")

            # Get PDFs from specific folders
            filter = FileFilter(
                folder_paths=["Documents/Reports", "Documents/Archive"],
                extensions=[".pdf"],
            )
            for file in client.list_files_filtered(filter):
                process_pdf(file)

            # Wildcard path matching
            filter = FileFilter(
                path_patterns=["Projects/2024-*/**/*.xlsx"],
            )
            for file in client.list_files_filtered(filter):
                print(file.get_full_path())
        """
        site_id = self.get_site_id()
        target_folders = file_filter.get_target_folders()

        if target_folders:
            # Search only specified folders
            for folder_path in target_folders:
                yield from self._walk_and_filter(
                    site_id,
                    file_filter,
                    folder_path=folder_path,
                    drive_id=drive_id,
                )
        else:
            # Search entire drive
            yield from self._walk_and_filter(
                site_id,
                file_filter,
                folder_path=None,
                drive_id=drive_id,
            )

    def _walk_and_filter(
        self,
        site_id: str,
        file_filter: FileFilter,
        *,
        folder_path: str | None = None,
        drive_id: str | None = None,
    ) -> Iterator[SharePointFileMetadata]:
        """
        Walk through a folder (or entire drive) and yield files matching the filter.

        Args:
            site_id: The SharePoint site ID
            file_filter: Filter criteria to apply
            folder_path: Optional folder path to start from (None = root)
            drive_id: Optional drive ID

        Yields:
            SharePointFileMetadata for matching files
        """
        if folder_path:
            # Get the folder item ID first, then walk from there
            folder_item = self._get_folder_by_path(site_id, folder_path, drive_id)
            if folder_item is None:
                logger.warning(f"Folder not found: {folder_path}")
                return
            item_id = folder_item.get("id")
            parent_path = folder_path
        else:
            item_id = None
            parent_path = ""

        for file_meta in self._walk_drive_items(
            site_id, item_id, drive_id=drive_id, parent_path=parent_path
        ):
            if file_filter.matches(file_meta):
                yield file_meta

    def _get_folder_by_path(
        self,
        site_id: str,
        folder_path: str,
        drive_id: str | None = None,
    ) -> dict[str, Any] | None:
        """
        Get folder metadata by path.

        Args:
            site_id: The SharePoint site ID
            folder_path: Path to the folder (relative to drive root)
            drive_id: Optional drive ID

        Returns:
            Folder item dict or None if not found
        """
        encoded_path = quote(folder_path.strip("/"), safe="/")

        if drive_id is None:
            url = f"{_GRAPH_API_BASE}/sites/{site_id}/drive/root:/{encoded_path}"
        else:
            url = f"{_GRAPH_API_BASE}/sites/{site_id}/drives/{drive_id}/root:/{encoded_path}"

        try:
            data = self._get_json(url)
            # Verify it's a folder
            if "folder" in data:
                return data
            return None
        except SharePointRequestError as exc:
            if exc.status_code == 404:
                return None
            raise

    def list_files_modified_since(
        self,
        since: datetime,
        *,
        folder_paths: list[str] | None = None,
        extensions: list[str] | None = None,
        drive_id: str | None = None,
    ) -> Iterator[SharePointFileMetadata]:
        """
        Convenience method for delta-sync: list files modified since a given datetime.

        This is a shorthand for creating a FileFilter with modified_after set.

        Args:
            since: Only include files modified on or after this datetime
            folder_paths: Optional list of folder paths to search
            extensions: Optional list of file extensions to filter (e.g., [".pdf", ".docx"])
            drive_id: Optional drive ID to search

        Yields:
            SharePointFileMetadata for each matching file

        Example:
            # Get all files modified in the last week
            from datetime import datetime, timedelta, timezone

            one_week_ago = datetime.now(timezone.utc) - timedelta(days=7)
            for file in client.list_files_modified_since(one_week_ago):
                print(f"{file.name} - modified: {file.last_modified}")
        """
        file_filter = FileFilter(
            modified_after=since,
            folder_paths=folder_paths or [],
            extensions=extensions or [],
        )
        yield from self.list_files_filtered(file_filter, drive_id=drive_id)

    def list_files_created_since(
        self,
        since: datetime,
        *,
        folder_paths: list[str] | None = None,
        extensions: list[str] | None = None,
        drive_id: str | None = None,
    ) -> Iterator[SharePointFileMetadata]:
        """
        Convenience method for delta-sync: list files created since a given datetime.

        Args:
            since: Only include files created on or after this datetime
            folder_paths: Optional list of folder paths to search
            extensions: Optional list of file extensions to filter
            drive_id: Optional drive ID to search

        Yields:
            SharePointFileMetadata for each matching file
        """
        file_filter = FileFilter(
            created_after=since,
            folder_paths=folder_paths or [],
            extensions=extensions or [],
        )
        yield from self.list_files_filtered(file_filter, drive_id=drive_id)

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
