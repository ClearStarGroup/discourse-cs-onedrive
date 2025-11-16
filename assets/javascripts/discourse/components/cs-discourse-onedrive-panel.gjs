import Component from "@glimmer/component";
import { action } from "@ember/object";
import { tracked } from "@glimmer/tracking";
import { service } from "@ember/service";
import { ajax } from "discourse/lib/ajax";
import { extractError } from "discourse/lib/ajax-error";
import loadScript from "discourse/lib/load-script";
import msalLib, {
  InteractionRequiredAuthError as MsalInteractionRequiredAuthError,
} from "discourse/plugins/cs-discourse-onedrive/discourse/lib/msal-browser";
import graphLib from "discourse/plugins/cs-discourse-onedrive/discourse/lib/microsoft-graph-client";
import DButton from "discourse/components/d-button";
import loadingSpinner from "discourse/helpers/loading-spinner";
import icon from "discourse/helpers/d-icon";
import { i18n } from "discourse-i18n";
import { getAbsoluteURL } from "discourse-common/lib/get-url";
const GRAPH_SCOPES = ["Files.Read.All"];
const REDIRECT_PATH = "/cs-discourse-onedrive/auth/callback";
let msalClient;
let msalLibPromise;
let graphLibPromise;
let InteractionRequiredAuthErrorClass;

function getRedirectUri() {
  return getAbsoluteURL(REDIRECT_PATH);
}

function loadMsalLibrary() {
  if (!msalLibPromise) {
    InteractionRequiredAuthErrorClass =
      MsalInteractionRequiredAuthError ||
      window.msal?.InteractionRequiredAuthError;
    msalLibPromise = Promise.resolve(msalLib);
  }

  return msalLibPromise;
}

async function ensureMsalClient(siteSettings) {
  const clientId = siteSettings.cs_discourse_onedrive_client_id;
  if (!clientId) {
    return null;
  }

  const tenantId = siteSettings.cs_discourse_onedrive_tenant_id;
  if (!tenantId || tenantId === "common") {
    console.error(
      "OneDrive tenant ID is not configured. Please set cs_discourse_onedrive_tenant_id in site settings to your Azure AD tenant ID (not 'common' unless your app is multi-tenant)."
    );
    // For single-tenant apps, we need the actual tenant ID
    // You can find it in Azure Portal > Azure Active Directory > Overview > Tenant ID
    throw new Error(
      "OneDrive tenant ID must be configured. Set it in site settings to your Azure AD tenant ID."
    );
  }

  const msalLib = await loadMsalLibrary();

  if (!msalClient) {
    msalClient = new msalLib.PublicClientApplication({
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
  }

  if (msalClient.initialize) {
    await msalClient.initialize();
  }

  return msalClient;
}

function loadGraphLibrary() {
  if (!graphLibPromise) {
    graphLibPromise = Promise.resolve(graphLib);
  }

  return graphLibPromise;
}

export default class CsDiscourseOnedrivePanel extends Component {
  @service currentUser;
  @service siteSettings;

  @tracked account = null;
  @tracked accessToken = null;
  @tracked loading = false;
  @tracked refreshing = false;
  @tracked linking = false;
  @tracked errorMessage = null;
  @tracked folder = this._initialFolder();
  @tracked files = null;
  @tracked hasLoadedFiles = false;

  constructor() {
    super(...arguments);
    this._initialize();
  }

  get topicId() {
    return this.args.post?.topic_id || this.args.post?.topicId;
  }

  get canManage() {
    return this.args.post?.cs_discourse_onedrive?.can_manage;
  }

  get folderLinked() {
    return this.folder && Object.keys(this.folder).length > 0;
  }

  get isEnabled() {
    return this.siteSettings.cs_discourse_onedrive_enabled;
  }

  get isConfigured() {
    return (
      this.siteSettings.cs_discourse_onedrive_client_id &&
      this.siteSettings.cs_discourse_onedrive_client_id.length > 0
    );
  }

  get shouldRender() {
    return this.isEnabled;
  }

  get signedInLabel() {
    if (!this.account) {
      return null;
    }

    const email =
      this.account.username || this.account.name || this.account.localAccountId;

    return i18n("cs_discourse_onedrive.signed_in_as", { email });
  }

  _initialFolder() {
    return (
      this.args.post?.cs_discourse_onedrive?.folder ||
      this.args.post?.topic?.cs_discourse_onedrive?.folder ||
      null
    );
  }

  async _handleRedirectResult(client) {
    try {
      const result = await client.handleRedirectPromise();
      if (result?.account) {
        this.account = result.account;
        if (result.accessToken) {
          this.accessToken = result.accessToken;
        }
      }
      return result;
    } catch (error) {
      this._surfaceError(error);
      return null;
    }
  }

  async _initialize() {
    if (!this.shouldRender || !this.isConfigured) {
      return;
    }

    const client = await ensureMsalClient(this.siteSettings);
    if (!client) {
      return;
    }

    await this._handleRedirectResult(client);

    const accounts = client.getAllAccounts();
    if (!this.account && accounts.length > 0) {
      this.account = accounts[0];
    }

    if (this.account && !this.accessToken) {
      try {
        await this._ensureAccessToken();
      } catch (error) {
        if (
          !InteractionRequiredAuthErrorClass ||
          !(error instanceof InteractionRequiredAuthErrorClass)
        ) {
          this._surfaceError(error);
        }
      }
    }

    if (this.folderLinked && this.account) {
      await this._loadFiles();
    }
  }

  async _ensureAccessToken({ interactive = false, usePopup = false } = {}) {
    const client = await ensureMsalClient(this.siteSettings);
    if (!client) {
      return null;
    }

    if (!this.account && !interactive) {
      return null;
    }

    if (this.account) {
      try {
        const result = await client.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: this.account,
        });
        this.account = result.account;
        this.accessToken = result.accessToken;
        return this.accessToken;
      } catch (error) {
        if (
          InteractionRequiredAuthErrorClass &&
          error instanceof InteractionRequiredAuthErrorClass &&
          interactive
        ) {
          if (usePopup) {
            const result = await client.acquireTokenPopup({
              scopes: GRAPH_SCOPES,
              account: this.account,
            });
            this.account = result.account;
            this.accessToken = result.accessToken;
            return this.accessToken;
          } else {
            await this._interactiveTokenAcquisition(client, usePopup);
            return null;
          }
        } else {
          throw error;
        }
      }
    }

    if (interactive) {
      const token = await this._interactiveTokenAcquisition(client, usePopup);
      return token;
    }

    return null;
  }

  async _interactiveTokenAcquisition(client, usePopup = false) {
    this.loading = true;
    try {
      if (usePopup) {
        const result = await client.loginPopup({
          scopes: GRAPH_SCOPES,
        });
        this.account = result.account;
        this.accessToken = result.accessToken;
        return this.accessToken;
      } else {
        const startPage = window.location.href;
        window.sessionStorage.setItem("csod.msal.redirect", startPage);
        await client.loginRedirect({
          scopes: GRAPH_SCOPES,
          redirectStartPage: startPage,
        });
      }
    } finally {
      this.loading = false;
    }
    return null;
  }

  async _loadFiles() {
    if (!this.folderLinked) {
      return;
    }

    this.refreshing = true;
    this.errorMessage = null;

    try {
      const token = await this._ensureAccessToken({ interactive: false });

      if (!token) {
        // User needs to sign in again.
        return;
      }

      const graphLib = await loadGraphLibrary();
      const client = graphLib.Client.init({
        authProvider: (done) => {
          done(null, token);
        },
      });

      const driveId = this.folder.drive_id || this.folder.driveId;
      const itemId = this.folder.item_id || this.folder.itemId;

      if (!driveId || !itemId) {
        return;
      }

      const response = await client
        .api(
          `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children`
        )
        .get();

      this.files = response?.value || [];
      this.hasLoadedFiles = true;
    } catch (error) {
      this._surfaceError(error, "cs_discourse_onedrive.refresh_error");
    } finally {
      this.refreshing = false;
    }
  }

  _surfaceError(error, translationKey = "cs_discourse_onedrive.auth_error") {
    const apiError = extractError(error);
    if (apiError && typeof apiError === "string") {
      this.errorMessage = apiError;
    } else {
      this.errorMessage = i18n(translationKey);
    }
  }

  async _openPicker() {
    console.log("_openPicker: Starting v8 implementation");
    console.log("_openPicker: Current accessToken:", !!this.accessToken);

    const client = await ensureMsalClient(this.siteSettings);
    if (!client) {
      throw new Error("MSAL client not available");
    }

    // Get the user's OneDrive URL first to determine the baseUrl
    // We need this to get the correct SharePoint token (not Graph token)
    const graphLib = await loadGraphLibrary();
    const graphClient = graphLib.Client.init({
      authProvider: (done) => {
        // Use existing token if available for Graph call to get baseUrl
        done(null, this.accessToken);
      },
    });

    let pickerBaseUrl;
    try {
      console.log("_openPicker: Getting user's OneDrive information");
      const driveResponse = await graphClient.api("/me/drive").get();
      const webUrl = driveResponse.webUrl;
      console.log("_openPicker: OneDrive webUrl:", webUrl);

      // Extract the base URL (e.g., https://contoso-my.sharepoint.com)
      const urlMatch = webUrl.match(/^(https:\/\/[^\/]+)/);
      if (urlMatch) {
        pickerBaseUrl = urlMatch[1];
      } else {
        throw new Error(
          "Could not determine picker base URL from OneDrive webUrl"
        );
      }
    } catch (error) {
      console.error("_openPicker: Error getting OneDrive info:", error);
      // Fallback: try to construct URL from tenant ID
      const tenantId = this.siteSettings.cs_discourse_onedrive_tenant_id;
      if (tenantId && tenantId !== "common") {
        // Try to get tenant name from Graph
        try {
          const orgResponse = await graphClient.api("/organization").get();
          const tenantName = orgResponse.value?.[0]?.displayName
            ?.toLowerCase()
            .replace(/\s+/g, "");
          if (tenantName) {
            pickerBaseUrl = `https://${tenantName}-my.sharepoint.com`;
          }
        } catch (orgError) {
          console.error(
            "_openPicker: Error getting organization info:",
            orgError
          );
        }
      }

      if (!pickerBaseUrl) {
        throw new Error(
          "Could not determine picker base URL. Please ensure you have access to OneDrive."
        );
      }
    }

    console.log("_openPicker: Picker baseUrl:", pickerBaseUrl);

    // IMPORTANT: The picker requires SharePoint tokens, not Graph tokens
    // The resource should be the baseUrl with .default scope
    // Reference: https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/?view=odsp-graph-online
    const sharePointScopes = [`${pickerBaseUrl}/.default`];

    // Get SharePoint token (not Graph token) for the picker
    let tokenResult;
    if (this.account) {
      try {
        tokenResult = await client.acquireTokenSilent({
          scopes: sharePointScopes,
          account: this.account,
        });
      } catch (error) {
        if (
          InteractionRequiredAuthErrorClass &&
          error instanceof InteractionRequiredAuthErrorClass
        ) {
          tokenResult = await client.acquireTokenPopup({
            scopes: sharePointScopes,
            account: this.account,
          });
        } else {
          throw error;
        }
      }
    } else {
      tokenResult = await client.loginPopup({
        scopes: sharePointScopes,
      });
    }

    this.account = tokenResult.account;
    // Store SharePoint token for picker
    const sharePointToken = tokenResult.accessToken;

    console.log("_openPicker: SharePoint token acquired:", !!sharePointToken);

    if (!sharePointToken) {
      throw new Error("No SharePoint access token available");
    }

    // Construct picker configuration based on v8 API
    // Reference: https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/?view=odsp-graph-online
    const channelId = Math.random().toString(36).substring(2, 15);
    // Use the full origin for messaging - this is where postMessage will be sent
    const origin = window.location.origin;

    console.log("_openPicker: Using origin for messaging:", origin);

    // According to the documentation, the entry should be { oneDrive: {} } for OneDrive
    // Configure to allow both files and folders, but we'll validate folders in the pick handler
    const filePickerParams = {
      sdk: "8.0",
      entry: {
        oneDrive: {},
      },
      authentication: {},
      messaging: {
        origin: origin,
        channelId: channelId,
      },
      typesAndSources: {
        mode: "folders", // Set mode to "folders" to allow folder selection
        pivots: {
          oneDrive: true,
          recent: true,
        },
        // Filter to only show folders
        filters: ["folder"],
      },
    };

    // Construct the picker URL using _layouts/15/FilePicker.aspx
    const queryString = new URLSearchParams({
      filePicker: JSON.stringify(filePickerParams),
    });

    const pickerUrl = `${pickerBaseUrl}/_layouts/15/FilePicker.aspx?${queryString.toString()}`;
    console.log("_openPicker: Picker URL:", pickerUrl);

    return new Promise((resolve, reject) => {
      // Store reference to this and variables for use in nested functions
      const self = this;
      const baseUrl = pickerBaseUrl; // Store for use in authenticate handler
      const msalClient = client; // Store for use in authenticate handler
      let port = null;
      let messageHandler = null;

      // Open picker in a popup window (empty initially)
      const pickerWindow = window.open(
        "",
        "OneDrivePicker",
        "width=800,height=600,resizable=yes,scrollbars=yes"
      );

      if (!pickerWindow) {
        reject(
          new Error("Failed to open picker window - popup may be blocked")
        );
        return;
      }

      // Add a timeout to prevent the button from staying disabled forever
      const timeout = setTimeout(() => {
        pickerWindow.close();
        if (port) {
          try {
            port.close();
          } catch (e) {
            // Ignore errors
          }
        }
        if (messageHandler) {
          window.removeEventListener("message", messageHandler);
        }
        reject(
          new Error("OneDrive picker timed out - no response after 60 seconds")
        );
      }, 60000);

      // Listen for postMessage from the picker (MessageChannel initialization)
      messageHandler = (event) => {
        console.log("_openPicker: Message event received:", {
          origin: event.origin,
          data: event.data,
          source: event.source,
          ports: event.ports?.length || 0,
        });

        // Verify origin for security
        if (
          !event.origin.includes("sharepoint.com") &&
          !event.origin.includes("onedrive.live.com")
        ) {
          console.warn(
            "_openPicker: Ignoring message from untrusted origin:",
            event.origin
          );
          return;
        }

        console.log("_openPicker: Received message from picker:", event.data);
        console.log("_openPicker: Message origin:", event.origin);

        const message = event.data;

        // Handle MessageChannel initialization
        if (
          message.type === "initialize" &&
          message.channelId === channelId &&
          event.ports &&
          event.ports.length > 0
        ) {
          console.log("_openPicker: Initializing MessageChannel");
          port = event.ports[0];

          port.addEventListener("message", (portMessage) => {
            const msg = portMessage.data;
            console.log("_openPicker: Port message received:", msg);

            switch (msg.type) {
              case "notification":
                const notification = msg.data;
                console.log("_openPicker: Notification:", notification);

                // The picker sends a "page-loaded" notification when ready
                if (notification.notification === "page-loaded") {
                  console.log("_openPicker: Picker page loaded and ready");
                }

                // Log selection changes if any
                if (notification.selection) {
                  console.log(
                    "_openPicker: Selection changed:",
                    notification.selection
                  );
                }
                break;

              case "command":
                // Acknowledge the command
                port.postMessage({
                  type: "acknowledge",
                  id: msg.id,
                });

                const command = msg.data;

                switch (command.command) {
                  case "authenticate":
                    // Return the SharePoint token for the requested resource
                    // The command contains the resource that needs authentication
                    console.log(
                      "_openPicker: Authentication requested for resource:",
                      command.resource
                    );

                    // Get a SharePoint token for the requested resource
                    // The resource should be the baseUrl
                    const authResource = command.resource || baseUrl;
                    const authScopes = [`${authResource}/.default`];

                    msalClient
                      .acquireTokenSilent({
                        scopes: authScopes,
                        account: self.account,
                      })
                      .then((authResult) => {
                        port.postMessage({
                          type: "result",
                          id: msg.id,
                          data: {
                            result: "token",
                            token: authResult.accessToken,
                          },
                        });
                      })
                      .catch((authError) => {
                        // Fallback to popup if silent fails
                        if (
                          InteractionRequiredAuthErrorClass &&
                          authError instanceof InteractionRequiredAuthErrorClass
                        ) {
                          msalClient
                            .acquireTokenPopup({
                              scopes: authScopes,
                              account: self.account,
                            })
                            .then((authResult) => {
                              port.postMessage({
                                type: "result",
                                id: msg.id,
                                data: {
                                  result: "token",
                                  token: authResult.accessToken,
                                },
                              });
                            })
                            .catch((popupError) => {
                              port.postMessage({
                                type: "result",
                                id: msg.id,
                                data: {
                                  result: "error",
                                  error: {
                                    code: "unableToObtainToken",
                                    message: popupError.message,
                                  },
                                },
                              });
                            });
                        } else {
                          port.postMessage({
                            type: "result",
                            id: msg.id,
                            data: {
                              result: "error",
                              error: {
                                code: "unableToObtainToken",
                                message: authError.message,
                              },
                            },
                          });
                        }
                      });
                    break;

                  case "close":
                    console.log("_openPicker: Close command received");
                    clearTimeout(timeout);
                    window.removeEventListener("message", messageHandler);
                    pickerWindow.close();
                    if (port) {
                      try {
                        port.close();
                      } catch (e) {
                        // Ignore errors
                      }
                    }
                    reject(new Error("cancelled"));
                    break;

                  case "pick":
                    console.log("_openPicker: Pick command received:", command);
                    console.log(
                      "_openPicker: Command items:",
                      JSON.stringify(command.items, null, 2)
                    );

                    // Extract selection from command
                    const items = command.items || [];
                    if (items.length === 0) {
                      port.postMessage({
                        type: "result",
                        id: msg.id,
                        data: {
                          result: "error",
                          error: {
                            code: "noSelection",
                            message: "No folder selected",
                          },
                        },
                      });
                      reject(
                        new Error(
                          i18n("cs_discourse_onedrive.picker_requires_folder")
                        )
                      );
                      return;
                    }

                    // Log the first item to see its structure
                    if (items.length > 0) {
                      console.log(
                        "_openPicker: First item structure:",
                        items[0]
                      );
                      console.log(
                        "_openPicker: First item keys:",
                        Object.keys(items[0])
                      );
                    }

                    // Since we're in "folders" mode, all items returned should be folders
                    // The v8 picker in "folders" mode will only return folders
                    // So we can safely use the first item
                    const selectedItem = items[0];

                    console.log("_openPicker: Selected item:", selectedItem);
                    console.log("_openPicker: Item type:", selectedItem.type);
                    console.log(
                      "_openPicker: Item has folder property:",
                      selectedItem.folder !== undefined
                    );

                    if (!selectedItem) {
                      port.postMessage({
                        type: "result",
                        id: msg.id,
                        data: {
                          result: "error",
                          error: {
                            code: "noSelection",
                            message: "No item selected",
                          },
                        },
                      });
                      reject(
                        new Error(
                          i18n("cs_discourse_onedrive.picker_requires_folder")
                        )
                      );
                      return;
                    }

                    // In "folders" mode, the picker only returns folders, so no need to validate
                    // Just proceed with the selected item

                    // Success - return the folder info
                    port.postMessage({
                      type: "result",
                      id: msg.id,
                      data: {
                        result: "success",
                      },
                    });

                    // Clean up
                    clearTimeout(timeout);
                    window.removeEventListener("message", messageHandler);
                    pickerWindow.close();
                    if (port) {
                      try {
                        port.close();
                      } catch (e) {
                        // Ignore errors
                      }
                    }

                    const parent =
                      selectedItem.parentReference || selectedItem.parent || {};

                    resolve({
                      drive_id: parent.driveId || parent.drive_id,
                      item_id: selectedItem.id || selectedItem.item_id,
                      name: selectedItem.name,
                      path: parent.path || parent.sharepointIds?.siteUrl,
                      web_url: selectedItem.webUrl || selectedItem.web_url,
                    });
                    break;

                  default:
                    console.warn(
                      "_openPicker: Unsupported command:",
                      command.command
                    );
                    port.postMessage({
                      type: "result",
                      id: msg.id,
                      data: {
                        result: "error",
                        error: {
                          code: "unsupportedCommand",
                          message: command.command,
                        },
                        isExpected: true,
                      },
                    });
                    break;
                }
                break;

              default:
                console.log("_openPicker: Unhandled message type:", msg.type);
                break;
            }
          });

          port.start();

          // Activate the picker
          port.postMessage({
            type: "activate",
          });
        }
      };

      window.addEventListener("message", messageHandler);

      // Wait for the window to be ready before creating and submitting the form
      // This ensures the document is available
      const submitForm = () => {
        try {
          // Ensure document is ready
          if (!pickerWindow.document) {
            console.log("_openPicker: Waiting for window document...");
            setTimeout(submitForm, 50);
            return;
          }

          // Write basic HTML structure if body doesn't exist yet
          if (!pickerWindow.document.body) {
            pickerWindow.document.write(
              "<!DOCTYPE html><html><head><title>OneDrive Picker</title></head><body></body></html>"
            );
            pickerWindow.document.close();
          }

          console.log(
            "_openPicker: Creating and submitting form with access token"
          );

          // Create and submit form with access token (POST method)
          // This is how v8 picker receives the access token
          const form = pickerWindow.document.createElement("form");
          form.setAttribute("action", pickerUrl);
          form.setAttribute("method", "POST");
          pickerWindow.document.body.appendChild(form);

          const input = pickerWindow.document.createElement("input");
          input.setAttribute("type", "hidden");
          input.setAttribute("name", "access_token");
          // Use SharePoint token, not Graph token
          input.setAttribute("value", sharePointToken);
          form.appendChild(input);

          console.log("_openPicker: Submitting form to:", pickerUrl);
          form.submit();
          console.log("_openPicker: Form submitted successfully");
        } catch (error) {
          console.error("_openPicker: Error submitting form:", error);
          reject(new Error(`Failed to submit picker form: ${error.message}`));
        }
      };

      // Start the form submission process after a brief delay to ensure window is ready
      setTimeout(submitForm, 100);

      // Handle window close as cancellation
      const checkClosed = setInterval(() => {
        if (pickerWindow.closed) {
          clearInterval(checkClosed);
          clearTimeout(timeout);
          window.removeEventListener("message", messageHandler);
          if (port) {
            try {
              port.close();
            } catch (e) {
              // Ignore errors
            }
          }
          reject(new Error("cancelled"));
        }
      }, 500);
    });
  }

  async _persistFolder(folder) {
    this.errorMessage = null;
    const url = `/cs-discourse-onedrive/topics/${this.topicId}/folder`;

    const response = await ajax(url, {
      type: "PUT",
      data: { folder },
    });

    this.folder = response.folder || null;
    this.files = null;
    this.hasLoadedFiles = false;
  }

  async _removeFolder() {
    this.errorMessage = null;
    const url = `/cs-discourse-onedrive/topics/${this.topicId}/folder`;

    const response = await ajax(url, {
      type: "DELETE",
    });

    this.folder = response.folder || null;
    this.files = null;
    this.hasLoadedFiles = false;
  }

  @action
  async signIn() {
    try {
      await this._ensureAccessToken({ interactive: true });
      if (this.folderLinked) {
        await this._loadFiles();
      }
    } catch (error) {
      this._surfaceError(error);
    }
  }

  @action
  async signOut() {
    const client = await ensureMsalClient(this.siteSettings);
    if (!client || !this.account) {
      return;
    }

    try {
      await client.logoutPopup({ account: this.account });
    } finally {
      this.account = null;
      this.accessToken = null;
      this.files = null;
      this.hasLoadedFiles = false;
    }
  }

  @action
  async clearCacheAndReauth() {
    // Clear MSAL cache to force re-authentication with correct tenant
    const client = await ensureMsalClient(this.siteSettings);
    if (client) {
      const accounts = client.getAllAccounts();
      accounts.forEach((account) => {
        client.removeAccount(account);
      });
      // Clear session storage
      sessionStorage.removeItem("msal.cache");
      sessionStorage.removeItem("msal.account.keys");
      sessionStorage.removeItem("msal.idtoken.keys");
    }
    this.account = null;
    this.accessToken = null;
    // Re-authenticate
    await this.signIn();
  }

  @action
  async refreshFiles() {
    await this._loadFiles();
  }

  @action
  async linkFolder() {
    // Use both console and alert for debugging
    try {
      console.log("linkFolder action called");
      console.log("canManage:", this.canManage);

      if (!this.canManage) {
        console.log("linkFolder: canManage is false, returning early");
        return;
      }

      console.log("linkFolder: Setting linking to true");
      this.linking = true;
      this.errorMessage = null; // Clear any previous errors

      console.log("linkFolder: Calling _openPicker");
      const selection = await this._openPicker();
      console.log("linkFolder: Picker returned selection:", selection);
      await this._persistFolder(selection);
      await this._loadFiles();
    } catch (error) {
      console.error("linkFolder: Error caught:", error);
      console.error("linkFolder: Error stack:", error?.stack);
      if (error?.message !== "cancelled") {
        const errorMsg = error?.message || String(error);
        console.error("linkFolder: Error message:", errorMsg);
        this._surfaceError(error, "cs_discourse_onedrive.picker_error");
      }
    } finally {
      console.log("linkFolder: Setting linking to false");
      this.linking = false;
    }
  }

  @action
  async changeFolder() {
    await this.linkFolder();
  }

  @action
  async removeFolder() {
    if (!this.canManage) {
      return;
    }

    if (!confirm(i18n("cs_discourse_onedrive.confirm_remove"))) {
      return;
    }

    try {
      await this._removeFolder();
    } catch (error) {
      this._surfaceError(error, "cs_discourse_onedrive.save_error");
    }
  }

  <template>
    {{#if this.shouldRender}}
      {{#unless this.isConfigured}}
        {{#if this.canManage}}
          <div class="post__row row">
            <div class="post__body topic-body clearfix">
              <div class="alert alert-error">
                {{i18n "cs_discourse_onedrive.not_configured"}}
              </div>
            </div>
          </div>
        {{/if}}
      {{else}}
        {{#if this.errorMessage}}
          <div class="post__row row">
            <div class="post__body topic-body clearfix">
              <div class="alert alert-error">
                {{this.errorMessage}}
              </div>
            </div>
          </div>
        {{/if}}

        {{#if this.account}}
          {{#if this.folderLinked}}
            <div class="post__row row">
              <div class="topic-avatar cs-onedrive-avatar">
                {{icon "cloud"}}
              </div>

              <div class="post__body topic-body clearfix">
                <div class="topic-meta-data">
                  <div class="names">
                    <span class="first username cs-onedrive-header">
                      {{i18n "cs_discourse_onedrive.section_header"}}
                    </span>
                  </div>

                  {{#if this.canManage}}
                    <div class="post-infos">
                      <div class="actions">
                        <DButton
                          @icon="arrows-rotate"
                          @title="cs_discourse_onedrive.refresh"
                          @action={{this.refreshFiles}}
                          @disabled={{this.refreshing}}
                          class="btn no-text btn-icon post-action-menu__refresh btn-flat"
                        />
                        <DButton
                          @icon="folder"
                          @title="cs_discourse_onedrive.change_folder"
                          @action={{this.changeFolder}}
                          @disabled={{this.linking}}
                          class="btn no-text btn-icon post-action-menu__change-folder btn-flat"
                        />
                        <DButton
                          @icon="trash-can"
                          @title="cs_discourse_onedrive.remove_folder"
                          @action={{this.removeFolder}}
                          class="btn no-text btn-icon post-action-menu__remove btn-flat btn-danger"
                        />
                        <DButton
                          @icon="right-from-bracket"
                          @title="cs_discourse_onedrive.sign_out"
                          @action={{this.signOut}}
                          @disabled={{this.loading}}
                          class="btn no-text btn-icon post-action-menu__logout btn-flat"
                        />
                      </div>
                    </div>
                  {{/if}}
                </div>

                <div class="post__regular regular post__contents contents">
                  <div class="cooked">
                    <p>
                      {{#if this.folder.web_url}}
                        <a
                          href={{this.folder.web_url}}
                          target="_blank"
                          rel="noopener"
                        >
                          {{this.folder.name}}
                        </a>
                      {{else}}
                        {{this.folder.name}}
                      {{/if}}
                    </p>
                    {{#if this.refreshing}}
                      <div class="cs-onedrive-panel__spinner">
                        {{loadingSpinner}}
                        <span>{{i18n
                            "cs_discourse_onedrive.fetching_files"
                          }}</span>
                      </div>
                    {{else if this.hasLoadedFiles}}
                      {{#if this.files.length}}
                        <ul class="cs-onedrive-panel__list">
                          {{#each this.files as |file|}}
                            <li>
                              {{#if file.webUrl}}
                                <a
                                  href={{file.webUrl}}
                                  target="_blank"
                                  rel="noopener"
                                >
                                  {{file.name}}
                                </a>
                              {{else}}
                                {{file.name}}
                              {{/if}}
                            </li>
                          {{/each}}
                        </ul>
                      {{else}}
                        <p>{{i18n "cs_discourse_onedrive.no_files"}}</p>
                      {{/if}}
                    {{else if this.files}}
                      <ul class="cs-onedrive-panel__list">
                        {{#each this.files as |file|}}
                          <li>
                            {{#if file.webUrl}}
                              <a
                                href={{file.webUrl}}
                                target="_blank"
                                rel="noopener"
                              >
                                {{file.name}}
                              </a>
                            {{else}}
                              {{file.name}}
                            {{/if}}
                          </li>
                        {{/each}}
                      </ul>
                    {{/if}}
                  </div>
                </div>
              </div>
            </div>
          {{else}}
            <div class="post__row row">
              <div class="topic-avatar cs-onedrive-avatar">
                {{icon "cloud"}}
              </div>

              <div class="post__body topic-body clearfix">
                <div class="topic-meta-data">
                  <div class="names">
                    <span class="first username cs-onedrive-header">
                      {{i18n "cs_discourse_onedrive.section_header"}}
                    </span>
                  </div>
                </div>

                <div class="post__regular regular post__contents contents">
                  <div class="cooked">
                    {{#if this.canManage}}
                      <p>{{i18n "cs_discourse_onedrive.folder_link_help"}}</p>
                      <DButton
                        @label="cs_discourse_onedrive.link_folder"
                        @action={{this.linkFolder}}
                        @disabled={{this.linking}}
                        class="btn btn-primary"
                      />
                    {{/if}}
                  </div>
                </div>
              </div>
            </div>
          {{/if}}
        {{else}}
          <div class="post__row row">
            <div class="topic-avatar">
              <svg
                class="fa d-icon d-icon-cloud svg-icon svg-string"
                aria-hidden="true"
                xmlns="http://www.w3.org/2000/svg"
              >
                <use href="#cloud"></use>
              </svg>
            </div>

            <div class="post__body topic-body clearfix">
              <div class="topic-meta-data">
                <div class="names">
                  <span class="first username">
                    {{i18n "cs_discourse_onedrive.sign_in_required"}}
                  </span>
                </div>
              </div>

              <div class="post__regular regular post__contents contents">
                <div class="cooked">
                  <DButton
                    @label="cs_discourse_onedrive.sign_in"
                    @action={{this.signIn}}
                    @disabled={{this.loading}}
                    class="btn btn-primary"
                  />
                </div>
              </div>
            </div>
          </div>
        {{/if}}
      {{/unless}}
    {{/if}}
  </template>
}
