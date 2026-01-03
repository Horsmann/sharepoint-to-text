"""
Exceptions for SharePoint REST access and Entra ID authentication.
"""


class SharePointError(Exception):
    """Base exception for SharePoint REST errors."""


class SharePointAuthError(SharePointError):
    """Raised when Entra ID authentication fails or returns an invalid token."""


class SharePointRequestError(SharePointError):
    """Raised when a SharePoint REST request fails."""

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None,
        body: str | None,
        url: str,
    ) -> None:
        super().__init__(message)
        self.status_code = status_code
        self.body = body
        self.url = url
