# SharePoint Graph API Setup Guide

This guide explains how to configure Microsoft Entra ID (Azure AD) and SharePoint permissions to use the `sharepoint_io` module.

## Prerequisites

- Access to Microsoft Azure Portal with admin permissions
- A SharePoint site you want to access
- Python environment with this package installed

## Step 1: Register an Application in Entra ID

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations**
2. Click **New registration**
3. Configure the application:
   - **Name**: Choose a descriptive name (e.g., "SharePoint File Reader")
   - **Supported account types**: Select "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank (not needed for app-only auth)
4. Click **Register**
5. Note the **Application (client) ID** - you'll need this later

## Step 2: Create a Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description (e.g., "SharePoint access")
4. Select an expiry period (recommended: 12-24 months)
5. Click **Add**
6. **Important**: Copy the secret **Value** immediately - it won't be shown again

## Step 3: Configure API Permissions

1. Go to **API permissions** > **Add a permission**
2. Select **Microsoft Graph**
3. Select **Application permissions** (not Delegated)
4. Add the following permission:
   - `Sites.Selected` - Allows access only to specific sites you explicitly grant
5. Click **Add permissions**
6. Click **Grant admin consent for [your organization]**
7. Verify the status shows "Granted for [organization]"

### Why Sites.Selected?

`Sites.Selected` is the most secure option because:
- It requires explicit per-site access grants
- The app cannot access any SharePoint site by default
- You control exactly which sites the app can read

Alternative (less secure): `Sites.Read.All` grants read access to all sites.

## Step 4: Grant Access to Specific SharePoint Sites

The `Sites.Selected` permission requires explicitly granting access to each site. Choose one of these methods:

### Option A: Using Microsoft Graph Explorer

1. Go to [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
2. Sign in with an admin account
3. First, get your site ID:
   ```
   GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}
   ```
   Example: `GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/ProjectDocs`

4. Grant permission to your app:
   ```
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Content-Type: application/json

   {
     "roles": ["read"],
     "grantedToIdentities": [{
       "application": {
         "id": "{your-app-client-id}",
         "displayName": "SharePoint File Reader"
       }
     }]
   }
   ```

### Option B: Using PnP PowerShell

```powershell
# Install PnP PowerShell if needed
Install-Module -Name PnP.PowerShell

# Connect to SharePoint
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/ProjectDocs" -Interactive

# Grant read permission
Grant-PnPAzureADAppSitePermission `
    -AppId "{your-app-client-id}" `
    -DisplayName "SharePoint File Reader" `
    -Permissions Read `
    -Site "https://contoso.sharepoint.com/sites/ProjectDocs"
```

### Option C: Using Azure CLI

```bash
# Get site ID
az rest --method get \
    --url "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/ProjectDocs" \
    --query id -o tsv

# Grant permission (replace {site-id} and {client-id})
az rest --method post \
    --url "https://graph.microsoft.com/v1.0/sites/{site-id}/permissions" \
    --body '{
        "roles": ["read"],
        "grantedToIdentities": [{
            "application": {
                "id": "{client-id}",
                "displayName": "SharePoint File Reader"
            }
        }]
    }'
```

## Step 5: Create Environment Configuration

Create a `.env` file in the project root with your credentials:

```env
sp_tenant_id=your-tenant-id-guid
sp_client_id=your-app-client-id-guid
sp_client_secret=your-client-secret-value
sp_site_url=https://contoso.sharepoint.com/sites/ProjectDocs
```

Where to find these values:
- **tenant_id**: Azure Portal > Microsoft Entra ID > Overview > **Tenant ID**
- **client_id**: Your app registration > Overview > **Application (client) ID**
- **client_secret**: The secret value you copied in Step 2
- **site_url**: The full URL to your SharePoint site

**Security Note**: Never commit `.env` files to version control. Add `.env` to your `.gitignore`.

## Step 6: Validate the Setup

Run the test setup script to verify everything is configured correctly:

```bash
python -m sharepoint2text.sharepoint_io.run_test_setup
```

### Expected Output

A successful run will display:

```
Token claims: {'aud': 'https://graph.microsoft.com', 'roles': ['Sites.Selected'], ...}

--- Site ID ---
Site ID: contoso.sharepoint.com,guid1,guid2

--- Document Libraries ---
  - Documents (id: drive-id-1)
  - Shared Documents (id: drive-id-2)

--- All Files ---
  - folder/document.docx (12345 bytes)
  - report.pdf (67890 bytes)
      CustomField1: value1
      CustomField2: value2

--- Downloading First File as JSON ---
  Saved: document.docx.json

--- Files Modified in Last 30 Days ---
  Found 5 file(s):
    - folder/document.docx (modified: 2024-01-15T10:30:00Z)
    ...
```

The script will:
1. Authenticate and display token claims
2. Resolve the SharePoint site ID
3. List all document libraries
4. List all files with their metadata and custom fields
5. Download the first file and save it as a JSON file with base64-encoded content
6. Demonstrate filtered queries (files modified recently, PDFs only)

## Troubleshooting

### "Unsupported app only token"

**Cause**: Using SharePoint REST API scope instead of Graph API.

**Solution**: Ensure your token scope is `https://graph.microsoft.com/.default` (this is the default).

### "Access denied" or 403 Forbidden

**Cause**: The app doesn't have permission to the specific SharePoint site.

**Solution**:
1. Verify you completed Step 4 to grant site-specific access
2. Check that the site URL in your `.env` matches the site you granted access to
3. Ensure admin consent was granted in Step 3

### "Invalid client secret"

**Cause**: The client secret has expired or was copied incorrectly.

**Solution**:
1. Go to your app registration > Certificates & secrets
2. Check if the secret has expired
3. Create a new secret and update your `.env` file

### "Site not found" or 404

**Cause**: The site URL is incorrect or the site doesn't exist.

**Solution**:
1. Verify the site URL by opening it in a browser
2. Check for typos in the URL
3. Ensure the URL format is correct: `https://{tenant}.sharepoint.com/sites/{site-name}`

### No files returned

**Cause**: The document library is empty or files are in a different drive.

**Solution**:
1. Check if files exist in the SharePoint site via the web interface
2. The script lists all available drives - verify which one contains your files
3. Files in the root "Documents" library should appear by default

## Using the Client Programmatically

After setup is validated, you can use the client in your code:

```python
import os
from datetime import datetime, timedelta, timezone

import dotenv

from sharepoint2text.sharepoint_io.client import (
    EntraIDAppCredentials,
    FileFilter,
    SharePointRestClient,
)

dotenv.load_dotenv()

credentials = EntraIDAppCredentials(
    tenant_id=os.environ["sp_tenant_id"],
    client_id=os.environ["sp_client_id"],
    client_secret=os.environ["sp_client_secret"],
)

client = SharePointRestClient(
    site_url=os.environ["sp_site_url"],
    credentials=credentials,
)

# List all files
for file in client.list_all_files():
    print(f"{file.name} - {file.size} bytes")

# Delta sync: get files modified in the last week
one_week_ago = datetime.now(timezone.utc) - timedelta(days=7)
for file in client.list_files_modified_since(one_week_ago):
    print(f"Modified: {file.name}")

# Advanced filtering
filter = FileFilter(
    modified_after=one_week_ago,
    folder_paths=["Documents/Reports"],
    extensions=[".pdf", ".docx"],
)
for file in client.list_files_filtered(filter):
    content = client.download_file(file.id)
    # Process file content...
```

## Security Best Practices

1. **Use Sites.Selected**: Only grant access to specific sites, not all sites
2. **Rotate secrets**: Set calendar reminders to rotate client secrets before expiry
3. **Least privilege**: Only request "read" permissions unless write access is needed
4. **Audit access**: Periodically review which apps have access to your SharePoint sites
5. **Secure credentials**: Use environment variables or secret managers, never hardcode credentials
