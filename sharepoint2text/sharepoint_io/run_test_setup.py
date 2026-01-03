"""
Runner script for testing SharePoint Graph API client.

Microsoft Graph API Setup
=========================

To use this client, you need to set up an Entra ID (Azure AD) app registration
with the correct permissions. Follow these steps:

1. Register an Application in Entra ID
   - Go to Azure Portal > Entra ID > App registrations > New registration
   - Name your application (e.g., "SharePoint File Reader")
   - Select "Accounts in this organizational directory only"
   - Click Register

2. Create a Client Secret
   - In your app registration, go to Certificates & secrets
   - Click "New client secret"
   - Add a description and select an expiry period
   - Copy the secret value immediately (it won't be shown again)

3. Configure API Permissions
   - Go to API permissions > Add a permission > Microsoft Graph
   - Select "Application permissions" (not delegated)
   - Add: Sites.Selected (allows access only to specific sites you grant)
   - Click "Grant admin consent for [your organization]"

4. Grant Access to Specific SharePoint Sites
   The Sites.Selected permission requires explicitly granting access to each site.
   Use the Microsoft Graph API or PowerShell to grant access:

   Using Graph API (POST request):
   ```
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["read"],  // or ["write"] for read/write access
     "grantedToIdentities": [{
       "application": {
         "id": "{your-app-client-id}",
         "displayName": "SharePoint File Reader"
       }
     }]
   }
   ```

   Using PnP PowerShell:
   ```powershell
   Grant-PnPAzureADAppSitePermission -AppId "{client-id}" -DisplayName "SharePoint File Reader" -Permissions Read -Site "{site-url}"
   ```

5. Create a .env File
   Create a .env file in the project root with these variables:
   ```
   sp_tenant_id=your-tenant-id-guid
   sp_client_id=your-app-client-id-guid
   sp_client_secret=your-client-secret-value
   sp_site_url=https://yourtenant.sharepoint.com/sites/yoursite
   ```

   - tenant_id: Found in Entra ID > Overview > Tenant ID
   - client_id: Found in your app registration > Overview > Application (client) ID
   - client_secret: The secret value you copied in step 2
   - site_url: The full URL to your SharePoint site

6. Run the Script
   ```bash
   python -m sharepoint2text.sharepoint_io.list_files_runner
   ```

Troubleshooting
---------------
- "Unsupported app only token": You're using SharePoint REST API scope instead of
  Graph API. Ensure scope is https://graph.microsoft.com/.default

- "Access denied" or 403: The app doesn't have permission to the site. Verify you
  completed step 4 to grant site-specific access.

- "Invalid client secret": The secret may have expired or was copied incorrectly.
  Create a new secret in step 2.
"""

import base64
import json
import os

import dotenv

from sharepoint2text.sharepoint_io.client import (
    EntraIDAppCredentials,
    SharePointFileMetadata,
    SharePointRestClient,
)
from sharepoint2text.sharepoint_io.exceptions import SharePointRequestError


def save_file_as_json(
    client: SharePointRestClient,
    file_meta: SharePointFileMetadata,
    output_dir: str = ".",
) -> str:
    """
    Download a file from SharePoint and save it as a JSON file with base64 content.

    The JSON file contains the file content as a base64-encoded string along with
    all metadata (both standard and custom fields).

    Args:
        client: SharePoint client to use for downloading
        file_meta: Metadata of the file to download
        output_dir: Directory to save the JSON file (default: current directory)

    Returns:
        Path to the created JSON file
    """
    # Download the file content
    file_bytes = client.download_file(file_meta.id)

    # Encode content as base64
    content_base64 = base64.b64encode(file_bytes).decode("ascii")

    # Build the JSON structure with all metadata
    json_data = {
        "file_content_base64": content_base64,
        "metadata": {
            "name": file_meta.name,
            "id": file_meta.id,
            "web_url": file_meta.web_url,
            "download_url": file_meta.download_url,
            "size": file_meta.size,
            "mime_type": file_meta.mime_type,
            "last_modified": file_meta.last_modified,
            "created": file_meta.created,
            "parent_path": file_meta.parent_path,
        },
        "custom_fields": file_meta.custom_fields or {},
    }

    # Create output filename (same name as original file + .json)
    output_filename = f"{file_meta.name}.json"
    output_path = os.path.join(output_dir, output_filename)

    # Write the JSON file
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, ensure_ascii=False)

    return output_path


def _get_required_env(key: str) -> str:
    value = os.getenv(key)
    if not value:
        raise ValueError(f"Missing required environment variable: {key}")
    return value


def _decode_jwt_payload(token: str) -> dict[str, object]:
    parts = token.split(".")
    if len(parts) < 2:
        raise ValueError("Invalid token format")
    payload_b64 = parts[1]
    padding = "=" * (-len(payload_b64) % 4)
    raw = base64.urlsafe_b64decode(payload_b64 + padding)
    return json.loads(raw.decode("utf-8", errors="replace"))


def _print_token_claims(token: str) -> None:
    payload = _decode_jwt_payload(token)
    claims = {
        "aud": payload.get("aud"),
        "roles": payload.get("roles"),
        "scp": payload.get("scp"),
        "tid": payload.get("tid"),
        "appid": payload.get("appid"),
    }
    print("Token claims:", claims)


if __name__ == "__main__":
    dotenv.load_dotenv()

    site_url = _get_required_env("sp_site_url")
    credentials = EntraIDAppCredentials(
        tenant_id=_get_required_env("sp_tenant_id"),
        client_id=_get_required_env("sp_client_id"),
        client_secret=_get_required_env("sp_client_secret"),
        # scope defaults to https://graph.microsoft.com/.default
    )
    client = SharePointRestClient(site_url=site_url, credentials=credentials)

    try:
        token = client.fetch_access_token()
        _print_token_claims(token)
    except SharePointRequestError as exc:
        print(f"Token request failed: {exc}")
        if exc.body:
            print(f"Token error body: {exc.body}")
        raise

    try:
        print("\n--- Site ID ---")
        site_id = client.get_site_id()
        print(f"Site ID: {site_id}")

        print("\n--- Document Libraries ---")
        drives = client.list_drives()
        for drive in drives:
            print(f"  - {drive.get('name')} (id: {drive.get('id')})")

        print("\n--- All Files ---")
        files = client.list_all_files()
        if not files:
            print("  No files found")
        for f in files:
            path = f"{f.parent_path}/{f.name}" if f.parent_path else f.name
            size_str = f" ({f.size} bytes)" if f.size else ""
            print(f"  - {path}{size_str}")
            if f.custom_fields:
                for key, value in f.custom_fields.items():
                    print(f"      {key}: {value}")

        # Download first file and save as JSON with base64 content and metadata
        if files:
            print("\n--- Downloading First File as JSON ---")
            first_file = files[0]
            output_path = save_file_as_json(client, first_file, output_dir=".")
            print(f"  Saved: {output_path}")

    except SharePointRequestError as exc:
        print(f"\nSharePoint request failed: {exc}")
        if exc.body:
            print(f"Error body: {exc.body}")
        raise
