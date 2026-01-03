"""
SharePoint REST helpers for downloading files with Entra ID authentication.
"""

from sharepoint2text.sharepoint_rest.client import (
    EntraIDAppCredentials,
    SharePointRestClient,
)
from sharepoint2text.sharepoint_rest.exceptions import (
    SharePointAuthError,
    SharePointError,
    SharePointRequestError,
)

__all__ = [
    "EntraIDAppCredentials",
    "SharePointRestClient",
    "SharePointAuthError",
    "SharePointError",
    "SharePointRequestError",
]
