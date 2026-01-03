import json
import unittest
from urllib.parse import parse_qs

from sharepoint2text.sharepoint_rest.client import (
    EntraIDAppCredentials,
    SharePointRestClient,
)
from sharepoint2text.sharepoint_rest.exceptions import (
    SharePointAuthError,
    SharePointRequestError,
)

tc = unittest.TestCase()


class FakeResponse:
    def __init__(self, status: int, body: bytes) -> None:
        self.status = status
        self._body = body

    def read(self) -> bytes:
        return self._body

    def getcode(self) -> int:
        return self.status

    def close(self) -> None:
        return None


def _build_credentials(scope: str = "https://contoso.sharepoint.com/.default"):
    return EntraIDAppCredentials(
        tenant_id="tenant-123",
        client_id="client-abc",
        client_secret="secret-xyz",
        scope=scope,
    )


def test_build_file_url_encodes_path() -> None:
    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
    )

    url = client.build_file_url(
        "/sites/demo/Shared Documents/My File.txt",
    )

    expected = (
        "https://contoso.sharepoint.com/sites/demo/_api/web/"
        "GetFileByServerRelativeUrl('/sites/demo/Shared%20Documents/"
        "My%20File.txt')/$value"
    )
    tc.assertEqual(url, expected)


def test_build_file_url_requires_leading_slash() -> None:
    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
    )

    with tc.assertRaises(ValueError):
        client.build_file_url("sites/demo/Shared Documents/My File.txt")


def test_fetch_access_token_success() -> None:
    captured = {}

    def fake_request(request, timeout=None):
        captured["request"] = request
        payload = {"access_token": "token-123"}
        return FakeResponse(200, json.dumps(payload).encode("utf-8"))

    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
        request_func=fake_request,
    )

    token = client.fetch_access_token()

    tc.assertEqual(token, "token-123")
    req = captured["request"]
    tc.assertEqual(
        req.full_url,
        "https://login.microsoftonline.com/tenant-123/oauth2/v2.0/token",
    )
    payload = parse_qs(req.data.decode("utf-8"))
    tc.assertEqual(payload["client_id"], ["client-abc"])
    tc.assertEqual(payload["client_secret"], ["secret-xyz"])
    tc.assertEqual(payload["scope"], ["https://contoso.sharepoint.com/.default"])
    tc.assertEqual(payload["grant_type"], ["client_credentials"])

    headers = {key.lower(): value for key, value in req.header_items()}
    tc.assertEqual(headers["content-type"], "application/x-www-form-urlencoded")
    tc.assertEqual(req.get_method(), "POST")


def test_fetch_access_token_missing_token_raises() -> None:
    def fake_request(request, timeout=None):
        payload = {"token_type": "Bearer"}
        return FakeResponse(200, json.dumps(payload).encode("utf-8"))

    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
        request_func=fake_request,
    )

    with tc.assertRaises(SharePointAuthError):
        client.fetch_access_token()


def test_download_file_success() -> None:
    requests = []

    def fake_request(request, timeout=None):
        requests.append(request)
        if "oauth2" in request.full_url:
            payload = {"access_token": "token-abc"}
            return FakeResponse(200, json.dumps(payload).encode("utf-8"))
        return FakeResponse(200, b"file-bytes")

    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
        request_func=fake_request,
    )

    data = client.download_file("/sites/demo/Shared Documents/file.txt")

    tc.assertEqual(data, b"file-bytes")
    tc.assertEqual(len(requests), 2)

    file_request = requests[1]
    headers = {key.lower(): value for key, value in file_request.header_items()}
    tc.assertEqual(headers["authorization"], "Bearer token-abc")
    tc.assertEqual(headers["accept"], "application/octet-stream")
    tc.assertEqual(file_request.get_method(), "GET")


def test_download_file_non_2xx_raises() -> None:
    def fake_request(request, timeout=None):
        if "oauth2" in request.full_url:
            payload = {"access_token": "token-abc"}
            return FakeResponse(200, json.dumps(payload).encode("utf-8"))
        return FakeResponse(404, b"not found")

    client = SharePointRestClient(
        "https://contoso.sharepoint.com/sites/demo",
        _build_credentials(),
        request_func=fake_request,
    )

    with tc.assertRaises(SharePointRequestError) as exc:
        client.download_file("/sites/demo/Shared Documents/file.txt")

    tc.assertEqual(exc.exception.status_code, 404)
