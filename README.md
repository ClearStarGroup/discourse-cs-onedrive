# Discourse OneDrive Plugin

Integrate OneDrive folders with Discourse topics. Link a OneDrive folder to any topic and display its files directly in the topic view.

## Installation

Follow the directions at [Install a Plugin](https://meta.discourse.org/t/install-a-plugin/19157) using this repository URL.

### Microsoft Azure Setup

1. Create an [Azure App Registration](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) for a Single Tenant App
2. Configure redirect URI: `https://your-discourse-site.com/cs-discourse-onedrive/auth/callback`
3. Add API permissions (Microsoft Graph):
   - `Files.Read.All` (Delegated) - Required for reading OneDrive files and folders
   - `Sites.Read.All` (Delegated) - Required for SharePoint/OneDrive file picker functionality
4. Grant admin consent for the permissions (if required by your organization)
5. Copy the Client ID to the plugin settings
6. Copy your Tenant ID to the plugin settings (see Azure Portal > Azure Active Directory > Overview)

## Configuration

After installation, configure the plugin in **Admin > Settings > Plugins**:

1. Enable the plugin by checking **cs_discourse_onedrive_enabled**
2. Set **cs_discourse_onedrive_client_id** to your Microsoft Azure App Registration Client ID
3. Set **cs_discourse_onedrive_tenant_id** to your Tenant ID
4. Set **cs_discourse_onedrive_sharepoint_base_url** to your SharePoint base URL (e.g., `https://contoso.sharepoint.com`)
5. Set **cs_discourse_onedrive_sharepoint_site_name** to your SharePoint site name (e.g., `AllCompany`)

## Features

- **Link OneDrive folders to topics** - Authorized users can link a OneDrive folder to any topic
- **Display folder contents** - View files from the linked OneDrive folder directly in the topic
- **User authentication** - Each viewer authenticates with their own Microsoft account via MSAL
- **Permission-based access** - Only users who can edit the topic's first post can link, change, or remove folders
- **Per-topic scope** - Each topic maintains its own folder association independently

## How It Works

1. The OneDrive panel appears after the first post in topics (only on the main post, not replies)
2. Viewers authenticate with their Microsoft account when first accessing the panel
3. Authorized users (topic owner, staff, or users with edit permissions) can link a OneDrive folder to the topic
4. The folder picker is constrained to the SharePoint site configured in settings, ensuring only folders from your organization's designated location can be selected
5. All viewers can see and access files from the linked folder using their own Microsoft credentials
6. Files are displayed with icons, names, sizes, and last modified dates

## Requirements

- Discourse 2.7.0 or higher
- Microsoft Azure App Registration with appropriate permissions

## Documentation

See contents of /docs for technical solution notes

## Backlog

[] Small action post update on setting, changing, or clearing linked folder
[] Support navigating into folders in panel
[] Support uploading files into panel
[] Support button to enable inserting link to file from linked folder in editor
