import "../../vendor/msal-browser.min";
import { extractError } from "discourse/lib/ajax-error";
import { getAbsoluteURL } from "discourse/lib/get-url";
import { i18n } from "discourse-i18n";

const REDIRECT_PATH = "/cs-discourse-onedrive/auth/callback";
const GRAPH_SCOPES = ["Files.Read.All"];

// Get MSAL global from window after vendor library loads
const msalGlobal = window.msal;

// Validate that MSAL library loaded correctly
if (!msalGlobal?.PublicClientApplication) {
  throw new Error("MSAL global not available after loading bundle");
}

// Export MSAL library and InteractionRequiredAuthError
export const msalLibrary = msalGlobal;
export const InteractionRequiredAuthError =
  msalGlobal.InteractionRequiredAuthError;

/**
 * Get the redirect URI for OAuth callbacks
 * @returns {string} The absolute redirect URI
 */
export function getRedirectUri() {
  return getAbsoluteURL(REDIRECT_PATH);
}

// Module-level state (singleton)
let msalClient = null;

/**
 * Get the MSAL client instance (creates and initializes if needed)
 * @param {Object} siteSettings - Site settings object
 * @returns {Promise<Object|null>} Promise resolving to MSAL client or null if not configured
 * @throws {Error} If tenant ID is not configured properly
 */
export async function getMsalClient(siteSettings) {
  // Return existing client if already created
  if (msalClient) {
    return msalClient;
  }

  // Check if client ID is configured
  const clientId = siteSettings.cs_discourse_onedrive_client_id;
  if (!clientId) {
    return null;
  }

  // Check if tenant ID is configured
  const tenantId = siteSettings.cs_discourse_onedrive_tenant_id;
  if (!tenantId || tenantId === "common") {
    throw new Error(
      "OneDrive tenant ID must be configured. Set it in site settings to your Azure AD tenant ID."
    );
  }

  // Create an MSAL client
  msalClient = new msalLibrary.PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: getRedirectUri(),
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: true,
    },
  });

  // Initialize the MSAL client
  await msalClient.initialize();

  return msalClient;
}

/**
 * Handle any pending redirect results from MSAL
 * This should be called after page load to process OAuth redirects
 * @param {Object} siteSettings - Site settings object
 * @returns {Promise<void>}
 */
export async function handleRedirectResult(siteSettings) {
  const client = await getMsalClient(siteSettings);
  if (!client) {
    return;
  }
  // Process any pending redirect results (clears "interaction in progress" state)
  await client.handleRedirectPromise();
}

/**
 * Get accounts from MSAL cache
 * @param {Object} siteSettings - Site settings object
 * @returns {Promise<Array>} Promise resolving to array of account objects
 */
export async function getAccounts(siteSettings) {
  const client = await getMsalClient(siteSettings);
  if (!client) {
    return [];
  }
  return client.getAllAccounts();
}

/**
 * Get scopes for the specified resource
 * @param {Object} siteSettings - Site settings object
 * @param {string} resource - Resource type: 'graph' or 'sharepoint'
 * @returns {Array<string>} Scopes array for the specified resource
 */
function getScopesForResource(siteSettings, resource) {
  if (resource === "sharepoint") {
    const sharePointBaseUrl =
      siteSettings.cs_discourse_onedrive_sharepoint_base_url;
    if (!sharePointBaseUrl || sharePointBaseUrl.trim() === "") {
      throw new Error(
        "SharePoint base URL is not configured. Please set cs_discourse_onedrive_sharepoint_base_url in site settings."
      );
    }
    return [`${sharePointBaseUrl}/.default`];
  }

  // Default to Graph API scopes
  return GRAPH_SCOPES;
}

/**
 * Acquire an access token, using silent acquisition if possible or interactive redirect if needed
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @param {boolean} [interactive=false] - Whether to allow interactive auth
 * @param {string} [resource='graph'] - Resource type: 'graph' for Graph API, 'sharepoint' for SharePoint
 * @returns {Promise<string|null>} Access token string, or null if no account and interactive not allowed
 * @throws {Error} If token acquisition fails (except InteractionRequiredAuthError when interactive is true)
 */
export async function acquireAccessToken(
  siteSettings,
  account,
  interactive = false,
  resource = "graph"
) {
  // Get the MSAL client
  const client = await getMsalClient(siteSettings);
  if (!client) {
    return null;
  }

  // Get scopes for the specified resource
  const scopes = getScopesForResource(siteSettings, resource);

  // Try silent token acquisition if account exists
  if (account) {
    try {
      const result = await client.acquireTokenSilent({
        scopes,
        account,
      });

      return result.accessToken;
    } catch (error) {
      // If silent acquisition fails with InteractionRequiredAuthError and interactive is allowed,
      // fall through to interactive auth below, otherwise throw the error
      if (!(error instanceof InteractionRequiredAuthError) || !interactive) {
        throw error;
      }
    }
  }

  // Perform interactive redirect authentication if no account or silent acquisition required interaction
  if (interactive) {
    // Store the current page in session storage to return to after authentication
    const currentPage = window.location.href;
    window.sessionStorage.setItem("csod.msal.redirect", currentPage);

    // Perform interactive redirect authentication (redirects away immediately)
    await client.loginRedirect({
      scopes,
      redirectStartPage: currentPage,
    });

    // Redundant return as we'll be redirected away immediately
    return null;
  }

  // No account and interactive not allowed
  return null;
}

/**
 * Sign out the current user
 * @param {Object} siteSettings - Site settings object
 * @param {Object} account - Account to sign out
 * @returns {Promise<void>}
 */
export async function signOut(siteSettings, account) {
  const client = await getMsalClient(siteSettings);
  if (!client || !account) {
    return;
  }

  await client.logoutRedirect({ account });
}

/**
 * Extract and format error message from an error object
 * @param {Error|Object} error - Error object
 * @param {string} translationKey - Translation key for default error message
 * @returns {string} Formatted error message
 */
export function surfaceError(error, translationKey) {
  const apiError = extractError(error);
  if (apiError && typeof apiError === "string") {
    return apiError;
  } else {
    return i18n(translationKey);
  }
}
