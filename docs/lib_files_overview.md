# Lib Files overview

## onedrive-auth-service.js

Authentication flow management service that handles all MSAL client operations, token acquisition, refresh, and error handling. Directly imports `msal-browser.min.js` from vendor assets (which attaches to `window.msal`), validates it loaded correctly, and exports `msalLibrary`, `InteractionRequiredAuthError`, and `getRedirectUri()` for OAuth redirect configuration. Exports `getMsalClient()` to get the MSAL client instance (creates and initializes if needed), `acquireAccessToken()` for acquiring tokens (handles both silent and interactive redirect flows), `getAccounts()` for retrieving accounts from MSAL cache, `signOut()` for user logout, and `surfaceError()` for consistent error handling. OAuth redirects are handled by the callback page (`callback.html.erb`), which processes redirect results and stores authentication state in MSAL's sessionStorage cache. This service abstracts all MSAL client access from components, providing a clean API for authentication operations. Depends on Discourse's error extraction utilities, i18n for translations, and `GRAPH_SCOPES` constant. Used by picker and API services, and the panel component.

## onedrive-api-service.js

OneDrive API operations service that handles file loading, folder persistence, and folder removal. Directly imports `microsoft-graph-client.js` from vendor assets (which attaches to `window.MicrosoftGraph`), validates it loaded correctly, and exports `graphLibrary` constant. Exports `loadFiles()` to fetch files and path from a linked folder (returns `{path: string, files: Array}`), `persistFolder()` to save folder metadata to the backend, and `removeFolder()` to delete the folder association. Depends on `onedrive-auth-service` for token management and error handling, and Discourse's AJAX utilities for backend communication. Used by the panel component for all file and folder operations.

## onedrive-picker-service.js

OneDrive folder picker integration service that opens the Microsoft file picker and returns the selected folder metadata. Exports `openPicker()` which takes site settings and account as parameters, handles SharePoint access token acquisition internally, and returns a promise resolving to a folder object with drive_id, item_id, name, path, and web_url. The picker is constrained to the SharePoint site configured in settings (sharepoint_base_url and sharepoint_site_name). Depends on `onedrive-auth-service` for token acquisition and MSAL client access. Used by the panel component's `linkFolder` and `changeFolder` actions.

## file-type-utils.js

Pure utility functions for file type handling with no dependencies on other plugin modules. Exports `getFileTypeIcon()` to return icon names for file types, `getFileTypeName()` to extract file type/extension, and `formatFileSize()` to format file sizes using Discourse's I18n utilities. Only depends on Discourse's I18n for human-readable size formatting. Used by the panel component template helpers for displaying file information.
