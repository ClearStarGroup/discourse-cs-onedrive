import { i18n } from "discourse-i18n";
import { acquireAccessToken } from "./onedrive-auth-service";

// Constants
const PICKER_TIMEOUT_MS = 60000; // 60 seconds
const WINDOW_MONITOR_INTERVAL_MS = 500;
const FORM_SUBMIT_DELAY_MS = 100;
const FORM_SUBMIT_RETRY_DELAY_MS = 50;
const FORM_SUBMIT_MAX_RETRIES = 100; // ~5 seconds max retry time

// Allowed domain suffixes for picker postMessage (subdomain-aware matching)
const ALLOWED_DOMAIN_SUFFIXES = [".sharepoint.com", ".onedrive.live.com"];

/**
 * Verify that the message origin is from an allowed SharePoint/OneDrive domain
 * @param {string} origin - Origin to verify
 * @returns {boolean} True if origin is allowed
 */
function isAllowedOrigin(origin) {
  try {
    const url = new URL(origin);
    const hostname = url.hostname.toLowerCase();
    // Check if hostname ends with any allowed domain suffix
    // This allows subdomains like "contoso.sharepoint.com" but prevents "evil-sharepoint.com"
    return ALLOWED_DOMAIN_SUFFIXES.some((suffix) => {
      return hostname.endsWith(suffix) && hostname.length > suffix.length;
    });
  } catch {
    return false;
  }
}

/**
 * Send error response to picker port
 * @param {MessagePort} port - MessageChannel port
 * @param {string|number} messageId - Message ID to respond to
 * @param {string} errorCode - Error code
 * @param {string} errorMessage - Error message
 * @param {boolean} isExpected - Whether error is expected
 */
function sendErrorResponse(
  port,
  messageId,
  errorCode,
  errorMessage,
  isExpected = false
) {
  port.postMessage({
    type: "result",
    id: messageId,
    data: {
      result: "error",
      error: {
        code: errorCode,
        message: errorMessage,
        ...(isExpected && { isExpected: true }),
      },
    },
  });
}

/**
 * Validate SharePoint site settings
 * @param {Object} siteSettings - Site settings object
 * @throws {Error} If settings are invalid
 */
function validateSiteSettings(siteSettings) {
  const sharePointBaseUrl =
    siteSettings.cs_discourse_onedrive_sharepoint_base_url;

  if (!sharePointBaseUrl || sharePointBaseUrl.trim() === "") {
    throw new Error(
      "SharePoint base URL is not configured. Please set cs_discourse_onedrive_sharepoint_base_url in site settings."
    );
  }

  const sharePointSiteName =
    siteSettings.cs_discourse_onedrive_sharepoint_site_name;

  if (!sharePointSiteName || sharePointSiteName.trim() === "") {
    throw new Error(
      "SharePoint site name is not configured. Please set cs_discourse_onedrive_sharepoint_site_name in site settings."
    );
  }
}

/**
 * Clean up picker resources (window, port, message handler, timeout)
 * @param {Window|null} pickerWindow - Picker popup window
 * @param {MessagePort|null} port - MessageChannel port
 * @param {Function|null} messageHandler - Message event handler
 * @param {number|null} timeout - Timeout ID
 */
function cleanup(pickerWindow, port, messageHandler, timeout) {
  if (timeout) {
    clearTimeout(timeout);
  }
  if (messageHandler) {
    window.removeEventListener("message", messageHandler);
  }
  if (pickerWindow) {
    pickerWindow.close();
  }
  if (port) {
    try {
      port.close();
    } catch {
      // Ignore errors
    }
  }
}

/**
 * Handle authenticate command from picker
 * @param {MessagePort} port - MessageChannel port
 * @param {Object} msg - Message object with id
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 */
function handleAuthenticateCommand(port, msg, siteSettings, account) {
  acquireAccessToken(
    siteSettings,
    account,
    true, // interactive
    "sharepoint" // resource
  )
    .then((token) => {
      if (!token) {
        throw new Error("No access token available");
      }
      port.postMessage({
        type: "result",
        id: msg.id,
        data: {
          result: "token",
          token,
        },
      });
    })
    .catch((authError) => {
      sendErrorResponse(port, msg.id, "unableToObtainToken", authError.message);
    });
}

/**
 * Handle pick command from picker (folder selection)
 * @param {MessagePort} port - MessageChannel port
 * @param {Object} msg - Message object with id
 * @param {Object} command - Command data with items array
 * @param {Function} cleanupFn - Cleanup function
 * @param {Function} resolve - Promise resolve function
 * @param {Function} reject - Promise reject function
 */
function handlePickCommand(port, msg, command, cleanupFn, resolve, reject) {
  const items = command.items || [];
  if (items.length === 0 || !items[0]) {
    sendErrorResponse(port, msg.id, "noSelection", "No folder selected");
    reject(new Error(i18n("cs_discourse_onedrive.picker_requires_folder")));
    return;
  }

  // In "folders" mode, the picker only returns folders
  const selectedItem = items[0];

  // Success - return the folder info
  port.postMessage({
    type: "result",
    id: msg.id,
    data: {
      result: "success",
    },
  });

  cleanupFn();

  const parent = selectedItem.parentReference || selectedItem.parent || {};

  resolve({
    drive_id: parent.driveId || parent.drive_id,
    item_id: selectedItem.id || selectedItem.item_id,
    name: selectedItem.name,
    path: parent.path || parent.sharepointIds?.siteUrl,
    web_url: selectedItem.webUrl || selectedItem.web_url,
  });
}

/**
 * Handle close command from picker (cancellation)
 * @param {Function} cleanupFn - Cleanup function
 * @param {Function} reject - Promise reject function
 */
function handleCloseCommand(cleanupFn, reject) {
  cleanupFn();
  reject(new Error("cancelled"));
}

/**
 * Handle unknown/unsupported command from picker
 * @param {MessagePort} port - MessageChannel port
 * @param {Object} msg - Message object with id
 * @param {Object} command - Command data
 */
function handleUnknownCommand(port, msg, command) {
  sendErrorResponse(port, msg.id, "unsupportedCommand", command.command, true);
}

/**
 * Create port message handler for MessageChannel communication
 * @param {MessagePort} port - MessageChannel port
 * @param {Function} cleanupFn - Cleanup function
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @param {Function} resolve - Promise resolve function
 * @param {Function} reject - Promise reject function
 * @returns {Function} Port message handler function
 */
function createPortMessageHandler(
  port,
  cleanupFn,
  siteSettings,
  account,
  resolve,
  reject
) {
  return (portMessage) => {
    const msg = portMessage.data;

    switch (msg.type) {
      case "notification":
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
            handleAuthenticateCommand(port, msg, siteSettings, account);
            break;

          case "close":
            handleCloseCommand(cleanupFn, reject);
            break;

          case "pick":
            handlePickCommand(port, msg, command, cleanupFn, resolve, reject);
            break;

          default:
            handleUnknownCommand(port, msg, command);
            break;
        }
        break;

      default:
        break;
    }
  };
}

/**
 * Create window message handler for postMessage communication
 * @param {string} channelId - Channel ID for picker communication
 * @param {Object} portRef - Reference object to store MessagePort
 * @param {Function} cleanupFn - Cleanup function
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @param {Function} resolve - Promise resolve function
 * @param {Function} reject - Promise reject function
 * @returns {Function} Window message handler function
 */
function createWindowMessageHandler(
  channelId,
  portRef,
  cleanupFn,
  siteSettings,
  account,
  resolve,
  reject
) {
  return (event) => {
    // Verify origin for security (exact domain matching)
    if (!isAllowedOrigin(event.origin)) {
      return;
    }

    const message = event.data;

    // Handle MessageChannel initialization
    // Guard against re-initialization if port already exists
    if (
      message.type === "initialize" &&
      message.channelId === channelId &&
      event.ports &&
      event.ports.length > 0 &&
      !portRef.current
    ) {
      portRef.current = event.ports[0];
      const port = portRef.current;

      const portHandler = createPortMessageHandler(
        port,
        cleanupFn,
        siteSettings,
        account,
        resolve,
        reject
      );

      port.addEventListener("message", portHandler);
      port.start();

      // Activate the picker
      port.postMessage({
        type: "activate",
      });
    }
  };
}

/**
 * Submit picker form with access token
 * @param {Window} pickerWindow - Picker popup window
 * @param {string} pickerUrl - URL for picker form submission
 * @param {string} token - Access token
 * @param {Function} cleanupFn - Cleanup function
 * @param {Function} reject - Promise reject function
 */
function submitPickerForm(pickerWindow, pickerUrl, token, cleanupFn, reject) {
  let retryCount = 0;

  const attemptSubmit = () => {
    try {
      if (!pickerWindow.document) {
        retryCount++;
        if (retryCount >= FORM_SUBMIT_MAX_RETRIES) {
          cleanupFn();
          reject(
            new Error(
              "Failed to submit picker form - window document not available"
            )
          );
          return;
        }
        setTimeout(attemptSubmit, FORM_SUBMIT_RETRY_DELAY_MS);
        return;
      }

      // Write basic HTML structure if body doesn't exist yet
      if (!pickerWindow.document.body) {
        pickerWindow.document.write(
          "<!DOCTYPE html><html><head><title>OneDrive Picker</title></head><body></body></html>"
        );
        pickerWindow.document.close();
      }

      // Create and submit form with access token (POST method)
      // This is how v8 picker receives the access token
      const form = pickerWindow.document.createElement("form");
      form.setAttribute("action", pickerUrl);
      form.setAttribute("method", "POST");
      pickerWindow.document.body.appendChild(form);

      const input = pickerWindow.document.createElement("input");
      input.setAttribute("type", "hidden");
      input.setAttribute("name", "access_token");
      input.setAttribute("value", token);
      form.appendChild(input);

      form.submit();
    } catch (error) {
      cleanupFn();
      reject(new Error(`Failed to submit picker form: ${error.message}`));
    }
  };

  // Start the form submission process after a brief delay to ensure window is ready
  setTimeout(attemptSubmit, FORM_SUBMIT_DELAY_MS);
}

/**
 * Monitor picker window for close events (cancellation)
 * @param {Window} pickerWindow - Picker popup window
 * @param {Function} cleanupFn - Cleanup function
 * @param {Function} reject - Promise reject function
 * @returns {number} Interval ID for cleanup
 */
function monitorPickerWindow(pickerWindow, cleanupFn, reject) {
  return setInterval(() => {
    if (pickerWindow.closed) {
      cleanupFn();
      reject(new Error("cancelled"));
    }
  }, WINDOW_MONITOR_INTERVAL_MS);
}

/**
 * Open the OneDrive folder picker and return the selected folder
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @returns {Promise<Object>} Promise resolving to selected folder object with drive_id, item_id, name, path, web_url
 * @throws {Error} If picker fails or is cancelled
 */
export async function openPicker(siteSettings, account) {
  // Validate required settings
  validateSiteSettings(siteSettings);

  const sharePointBaseUrl =
    siteSettings.cs_discourse_onedrive_sharepoint_base_url;
  const sharePointSiteName =
    siteSettings.cs_discourse_onedrive_sharepoint_site_name;

  // Construct the full SharePoint site URL
  const sharePointSiteUrl = `${sharePointBaseUrl}/sites/${sharePointSiteName}`;

  // Get SharePoint access token
  const token = await acquireAccessToken(
    siteSettings,
    account,
    true, // interactive
    "sharepoint" // resource
  );

  if (!token) {
    throw new Error("No SharePoint access token available");
  }

  // Construct picker configuration based on v8 API
  // Reference: https://learn.microsoft.com/en-us/onedrive/developer/controls/file-pickers/?view=odsp-graph-online
  const origin = window.location.origin;
  const channelId = Math.random().toString(36).substring(2, 15);
  const filePickerParams = {
    sdk: "8.0",
    messaging: {
      origin,
      channelId,
    },
    // Start at the specified SharePoint site
    entry: {
      sharePoint: {
        byPath: {
          web: `${sharePointSiteUrl}`,
        },
      },
    },
    // We'll handle authentication ourselves
    authentication: {},
    // Only show folders from the specified SharePoint site
    typesAndSources: {
      mode: "folders",
      filters: ["folder"],
      locations: {
        sharePoint: {
          byPath: {
            web: `${sharePointSiteUrl}`,
          },
        },
      },
      // Disable all pivots
      pivots: {
        oneDrive: false,
        recent: false,
        shared: false,
        sharedLibraries: false,
        myOrganization: false,
        site: false,
      },
    },
  };

  // Construct the picker URL using _layouts/15/FilePicker.aspx
  const queryString = new URLSearchParams({
    filePicker: JSON.stringify(filePickerParams),
  });
  const pickerUrl = `${sharePointBaseUrl}/_layouts/15/FilePicker.aspx?${queryString.toString()}`;

  return new Promise((resolve, reject) => {
    // Use reference objects for values that may be mutated
    const portRef = { current: null };
    const intervalRef = { current: null };
    let messageHandler = null;
    let timeout = null;

    // Open picker in a popup window (empty initially)
    const pickerWindow = window.open(
      "",
      "OneDrivePicker",
      "width=800,height=600,resizable=yes,scrollbars=yes"
    );

    if (!pickerWindow) {
      reject(new Error("Failed to open picker window - popup may be blocked"));
      return;
    }

    // Create cleanup function with access to all resources
    const cleanupFn = () => {
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }
      cleanup(pickerWindow, portRef.current, messageHandler, timeout);
    };

    // Add a timeout to prevent the button from staying disabled forever
    timeout = setTimeout(() => {
      cleanupFn();
      reject(
        new Error(
          `OneDrive picker timed out - no response after ${PICKER_TIMEOUT_MS / 1000} seconds`
        )
      );
    }, PICKER_TIMEOUT_MS);

    // Create window message handler
    messageHandler = createWindowMessageHandler(
      channelId,
      portRef,
      cleanupFn,
      siteSettings,
      account,
      resolve,
      reject
    );

    window.addEventListener("message", messageHandler);

    // Submit picker form
    submitPickerForm(pickerWindow, pickerUrl, token, cleanupFn, reject);

    // Monitor window for close events
    intervalRef.current = monitorPickerWindow(pickerWindow, cleanupFn, reject);
  });
}
