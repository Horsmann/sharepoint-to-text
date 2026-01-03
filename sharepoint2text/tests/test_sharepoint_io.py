"""
Unit tests for the SharePoint Graph API client.

These tests mock HTTP calls so no actual SharePoint connection is needed.
"""

import io
import json
from unittest.mock import MagicMock

import pytest

from sharepoint2text.sharepoint_io import (
    EntraIDAppCredentials,
    SharePointAuthError,
    SharePointFileMetadata,
    SharePointRequestError,
    SharePointRestClient,
)


# Test fixtures
@pytest.fixture
def credentials():
    """Create test credentials."""
    return EntraIDAppCredentials(
        tenant_id="test-tenant-id",
        client_id="test-client-id",
        client_secret="test-client-secret",
    )


@pytest.fixture
def site_url():
    """Test SharePoint site URL."""
    return "https://contoso.sharepoint.com/sites/testsite"


def _make_mock_response(data: dict | bytes, status: int = 200) -> MagicMock:
    """Create a mock HTTP response."""
    mock_response = MagicMock()
    mock_response.status = status
    mock_response.getcode.return_value = status
    if isinstance(data, bytes):
        mock_response.read.return_value = data
    else:
        mock_response.read.return_value = json.dumps(data).encode("utf-8")
    return mock_response


class TestEntraIDAppCredentials:
    """Tests for EntraIDAppCredentials dataclass."""

    def test_default_scope(self):
        """Test that default scope is Graph API."""
        creds = EntraIDAppCredentials(
            tenant_id="tenant",
            client_id="client",
            client_secret="secret",
        )
        assert creds.scope == "https://graph.microsoft.com/.default"

    def test_custom_scope(self):
        """Test custom scope can be provided."""
        creds = EntraIDAppCredentials(
            tenant_id="tenant",
            client_id="client",
            client_secret="secret",
            scope="https://custom.scope/.default",
        )
        assert creds.scope == "https://custom.scope/.default"


class TestSharePointRestClientTokenFetch:
    """Tests for token fetching."""

    def test_fetch_access_token_success(self, credentials, site_url):
        """Test successful token fetch."""
        token_response = {"access_token": "test-token-12345"}
        mock_response = _make_mock_response(token_response)

        def mock_request(request, timeout):
            assert "login.microsoftonline.com" in request.full_url
            assert "test-tenant-id" in request.full_url
            return mock_response

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        token = client.fetch_access_token()
        assert token == "test-token-12345"

    def test_fetch_access_token_missing_token(self, credentials, site_url):
        """Test error when response lacks access_token."""
        token_response = {"error": "invalid_grant"}
        mock_response = _make_mock_response(token_response)

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=lambda r, timeout: mock_response,
        )
        with pytest.raises(SharePointAuthError, match="missing access_token"):
            client.fetch_access_token()

    def test_fetch_access_token_invalid_json(self, credentials, site_url):
        """Test error when response is not valid JSON."""
        mock_response = _make_mock_response(b"not json")

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=lambda r, timeout: mock_response,
        )
        with pytest.raises(SharePointAuthError, match="Invalid token response JSON"):
            client.fetch_access_token()


class TestSharePointRestClientSiteId:
    """Tests for site ID resolution."""

    def test_get_site_id_with_path(self, credentials):
        """Test getting site ID from URL with path."""
        site_url = "https://contoso.sharepoint.com/sites/testsite"
        site_response = {
            "id": "contoso.sharepoint.com,guid1,guid2",
            "displayName": "Test Site",
        }

        call_count = [0]

        def mock_request(request, timeout):
            call_count[0] += 1
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            # Should be: /sites/contoso.sharepoint.com:/sites/testsite
            assert "graph.microsoft.com" in request.full_url
            assert "contoso.sharepoint.com:/sites/testsite" in request.full_url
            return _make_mock_response(site_response)

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        site_id = client.get_site_id()
        assert site_id == "contoso.sharepoint.com,guid1,guid2"

    def test_get_site_id_cached(self, credentials, site_url):
        """Test that site ID is cached after first fetch."""
        call_count = [0]

        def mock_request(request, timeout):
            call_count[0] += 1
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            return _make_mock_response({"id": "cached-site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )

        # First call
        site_id1 = client.get_site_id()
        initial_calls = call_count[0]

        # Second call should use cache
        site_id2 = client.get_site_id()
        assert site_id1 == site_id2
        assert call_count[0] == initial_calls  # No additional calls


class TestSharePointRestClientListDrives:
    """Tests for listing document libraries."""

    def test_list_drives(self, credentials, site_url):
        """Test listing drives/document libraries."""
        drives_response = {
            "value": [
                {"id": "drive1", "name": "Documents"},
                {"id": "drive2", "name": "Shared Documents"},
            ]
        }

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/sites/" in request.full_url and "/drives" in request.full_url:
                if "drive" in request.full_url and "children" not in request.full_url:
                    return _make_mock_response(drives_response)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        drives = client.list_drives()
        assert len(drives) == 2
        assert drives[0]["name"] == "Documents"
        assert drives[1]["name"] == "Shared Documents"


class TestSharePointRestClientListFiles:
    """Tests for listing files."""

    def test_list_all_files_empty(self, credentials, site_url):
        """Test listing files when no files exist."""

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/children" in request.full_url:
                return _make_mock_response({"value": []})
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        files = client.list_all_files()
        assert files == []

    def test_list_all_files_with_files(self, credentials, site_url):
        """Test listing files with file items."""
        files_response = {
            "value": [
                {
                    "id": "file1",
                    "name": "document.pdf",
                    "file": {"mimeType": "application/pdf"},
                    "size": 12345,
                    "webUrl": "https://contoso.sharepoint.com/document.pdf",
                    "createdDateTime": "2024-01-01T00:00:00Z",
                    "lastModifiedDateTime": "2024-01-02T00:00:00Z",
                },
                {
                    "id": "file2",
                    "name": "spreadsheet.xlsx",
                    "file": {
                        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    },
                    "size": 54321,
                    "webUrl": "https://contoso.sharepoint.com/spreadsheet.xlsx",
                },
            ]
        }

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/children" in request.full_url:
                return _make_mock_response(files_response)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        files = client.list_all_files()
        assert len(files) == 2
        assert files[0].name == "document.pdf"
        assert files[0].size == 12345
        assert files[0].mime_type == "application/pdf"
        assert files[1].name == "spreadsheet.xlsx"

    def test_list_files_with_folders(self, credentials, site_url):
        """Test that folders are traversed and files inside are returned."""
        root_response = {
            "value": [
                {
                    "id": "folder1",
                    "name": "Folder A",
                    "folder": {"childCount": 1},
                },
                {
                    "id": "file1",
                    "name": "root-file.txt",
                    "file": {"mimeType": "text/plain"},
                    "size": 100,
                    "webUrl": "https://contoso.sharepoint.com/root-file.txt",
                },
            ]
        }
        folder_response = {
            "value": [
                {
                    "id": "file2",
                    "name": "nested-file.pdf",
                    "file": {"mimeType": "application/pdf"},
                    "size": 200,
                    "webUrl": "https://contoso.sharepoint.com/Folder A/nested-file.pdf",
                },
            ]
        }

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/items/folder1/children" in request.full_url:
                return _make_mock_response(folder_response)
            if "/children" in request.full_url:
                return _make_mock_response(root_response)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        files = client.list_all_files()
        assert len(files) == 2
        names = {f.name for f in files}
        assert "root-file.txt" in names
        assert "nested-file.pdf" in names

        # Check parent path is set for nested file
        nested = next(f for f in files if f.name == "nested-file.pdf")
        assert nested.parent_path == "Folder A"

    def test_list_files_with_custom_fields(self, credentials, site_url):
        """Test that custom fields are extracted from listItem.fields."""
        files_response = {
            "value": [
                {
                    "id": "file1",
                    "name": "report.pdf",
                    "file": {"mimeType": "application/pdf"},
                    "size": 1000,
                    "webUrl": "https://contoso.sharepoint.com/report.pdf",
                    "listItem": {
                        "fields": {
                            "id": "1",
                            "Title": "Report Title",
                            "Created": "2024-01-01",
                            "Modified": "2024-01-02",
                            "CustomCategory": "Finance",
                            "ApprovalStatus": "Approved",
                            "Priority": 1,
                        }
                    },
                },
            ]
        }

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/children" in request.full_url:
                return _make_mock_response(files_response)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        files = client.list_all_files()
        assert len(files) == 1
        f = files[0]
        assert f.custom_fields is not None
        # System fields should be filtered out
        assert "id" not in f.custom_fields
        assert "Title" not in f.custom_fields
        assert "Created" not in f.custom_fields
        # Custom fields should be present
        assert f.custom_fields["CustomCategory"] == "Finance"
        assert f.custom_fields["ApprovalStatus"] == "Approved"
        assert f.custom_fields["Priority"] == 1

    def test_list_files_pagination(self, credentials, site_url):
        """Test that pagination is handled correctly."""
        page1 = {
            "value": [
                {
                    "id": "file1",
                    "name": "file1.txt",
                    "file": {"mimeType": "text/plain"},
                    "size": 100,
                    "webUrl": "https://contoso.sharepoint.com/file1.txt",
                }
            ],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page",
        }
        page2 = {
            "value": [
                {
                    "id": "file2",
                    "name": "file2.txt",
                    "file": {"mimeType": "text/plain"},
                    "size": 200,
                    "webUrl": "https://contoso.sharepoint.com/file2.txt",
                }
            ],
        }

        request_urls = []

        def mock_request(request, timeout):
            request_urls.append(request.full_url)
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "next-page" in request.full_url:
                return _make_mock_response(page2)
            if "/children" in request.full_url:
                return _make_mock_response(page1)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        files = client.list_all_files()
        assert len(files) == 2
        assert files[0].name == "file1.txt"
        assert files[1].name == "file2.txt"
        # Verify pagination was followed
        assert any("next-page" in url for url in request_urls)


class TestSharePointRestClientDownload:
    """Tests for file download."""

    def test_download_file_by_id(self, credentials, site_url):
        """Test downloading a file by ID."""
        file_content = b"Hello, World! This is test content."

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/items/file123/content" in request.full_url:
                return _make_mock_response(file_content)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        content = client.download_file("file123")
        assert content == file_content

    def test_download_file_by_path(self, credentials, site_url):
        """Test downloading a file by path."""
        file_content = b"PDF content here..."

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            if "/root:/Documents/report.pdf:/content" in request.full_url:
                return _make_mock_response(file_content)
            return _make_mock_response({"id": "site-id"})

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        content = client.download_file_by_path("/Documents/report.pdf")
        assert content == file_content


class TestSharePointRestClientErrorHandling:
    """Tests for error handling."""

    def test_http_error_raises_request_error(self, credentials, site_url):
        """Test that HTTP errors are converted to SharePointRequestError."""
        from urllib.error import HTTPError

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            error = HTTPError(
                request.full_url,
                403,
                "Forbidden",
                {},
                io.BytesIO(b'{"error": "access_denied"}'),
            )
            raise error

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        with pytest.raises(SharePointRequestError) as exc_info:
            client.get_site_id()

        assert exc_info.value.status_code == 403
        assert "access_denied" in exc_info.value.body

    def test_network_error_raises_request_error(self, credentials, site_url):
        """Test that network errors are converted to SharePointRequestError."""
        from urllib.error import URLError

        def mock_request(request, timeout):
            if "login.microsoftonline.com" in request.full_url:
                return _make_mock_response({"access_token": "token"})
            raise URLError("Connection refused")

        client = SharePointRestClient(
            site_url=site_url,
            credentials=credentials,
            request_func=mock_request,
        )
        with pytest.raises(SharePointRequestError) as exc_info:
            client.get_site_id()

        assert exc_info.value.status_code is None
        assert "network error" in str(exc_info.value).lower()


class TestSharePointFileMetadata:
    """Tests for SharePointFileMetadata dataclass."""

    def test_metadata_with_all_fields(self):
        """Test creating metadata with all fields."""
        meta = SharePointFileMetadata(
            name="test.pdf",
            id="file-123",
            web_url="https://example.com/test.pdf",
            download_url="https://example.com/download/test.pdf",
            size=12345,
            mime_type="application/pdf",
            last_modified="2024-01-01T00:00:00Z",
            created="2023-12-01T00:00:00Z",
            parent_path="Documents/Reports",
            custom_fields={"Category": "Finance", "Year": 2024},
        )
        assert meta.name == "test.pdf"
        assert meta.size == 12345
        assert meta.custom_fields["Year"] == 2024

    def test_metadata_with_minimal_fields(self):
        """Test creating metadata with only required fields."""
        meta = SharePointFileMetadata(
            name="test.txt",
            id="file-456",
            web_url="https://example.com/test.txt",
        )
        assert meta.name == "test.txt"
        assert meta.size is None
        assert meta.custom_fields is None
        assert meta.parent_path is None
