"""
SharePoint REST client that downloads files using Entra ID app authentication.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Callable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from sharepoint2text.sharepoint_rest.exceptions import (
    SharePointAuthError,
    SharePointRequestError,
)

_TOKEN_ENDPOINT_TEMPLATE = (
    "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
)


@dataclass(frozen=True)
class EntraIDAppCredentials:
    """Client credentials for an Entra ID application."""

    tenant_id: str
    client_id: str
    client_secret: str
    scope: str


class SharePointRestClient:
    """Small SharePoint REST client for downloading file bytes."""

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

    def build_file_url(self, server_relative_url: str) -> str:
        """Build the SharePoint REST URL for a server-relative path."""
        if not server_relative_url.startswith("/"):
            raise ValueError("server_relative_url must start with '/'")
        encoded = quote(server_relative_url, safe="/")
        return (
            f"{self._site_url}/_api/web/GetFileByServerRelativeUrl('{encoded}')/$value"
        )

    def fetch_access_token(self) -> str:
        """Request an app-only access token from Entra ID."""
        token_url = _TOKEN_ENDPOINT_TEMPLATE.format(
            tenant_id=self._credentials.tenant_id
        )
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
        return access_token

    def download_file(self, server_relative_url: str) -> bytes:
        """Download a file by server-relative URL and return its bytes."""
        token = self.fetch_access_token()
        file_url = self.build_file_url(server_relative_url)
        request = Request(
            file_url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/octet-stream",
            },
            method="GET",
        )
        _, body = self._send(request, request_kind="file")
        return body

    def _send(self, request: Request, *, request_kind: str) -> tuple[int, bytes]:
        response = None
        try:
            response = self._request(request, timeout=self._timeout)
        except HTTPError as exc:
            body = exc.read()
            raise SharePointRequestError(
                f"SharePoint {request_kind} request failed with status {exc.code}",
                status_code=exc.code,
                body=body.decode("utf-8", errors="replace") if body else None,
                url=request.full_url,
            ) from exc
        except URLError as exc:
            raise SharePointRequestError(
                f"SharePoint {request_kind} request failed due to network error",
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
                f"SharePoint {request_kind} request returned status {status}",
                status_code=status,
                body=body.decode("utf-8", errors="replace") if body else None,
                url=request.full_url,
            )
        return status, body
