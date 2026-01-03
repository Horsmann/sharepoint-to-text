"""
SharePoint client using Microsoft Graph API with Entra ID app authentication.
"""

from sharepoint2text.sharepoint_rest.client import (
    EntraIDAppCredentials,
    SharePointFileMetadata,
    SharePointRestClient,
)
from sharepoint2text.sharepoint_rest.exceptions import (
    SharePointAuthError,
    SharePointError,
    SharePointRequestError,
)

__all__ = [
    "EntraIDAppCredentials",
    "SharePointFileMetadata",
    "SharePointRestClient",
    "SharePointAuthError",
    "SharePointError",
    "SharePointRequestError",
]
