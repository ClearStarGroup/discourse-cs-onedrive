import "../../vendor/microsoft-graph-client";
import { ajax } from "discourse/lib/ajax";
import { acquireAccessToken, surfaceError } from "./onedrive-auth-service";

// Get Graph global from window after vendor library loads
const graphGlobal = window.MicrosoftGraph;

// Validate that Graph library loaded correctly
if (!graphGlobal?.Client) {
  throw new Error("Microsoft Graph global not available after loading bundle");
}

// Export Graph library
export const graphLibrary = graphGlobal;

/**
 * Get a Graph client instance with authentication
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @returns {Promise<Object>} Promise resolving to Graph client instance
 * @throws {Error} If token acquisition fails
 */
async function getGraphClient(siteSettings, account) {
  const token = await acquireAccessToken(siteSettings, account, false);

  if (!token) {
    throw new Error("Unable to acquire access token");
  }

  return graphLibrary.Client.init({
    authProvider: (done) => {
      done(null, token);
    },
  });
}

/**
 * Load files and path from a linked OneDrive folder
 * @param {Object} folder - Folder object with drive_id and item_id
 * @param {Object} siteSettings - Site settings object
 * @param {Object|null} account - Current user account
 * @returns {Promise<{path: string, files: Array}>} Promise resolving to object with path and files array
 */
export async function loadFiles(folder, siteSettings, account) {
  if (!folder || (!folder.drive_id && !folder.driveId)) {
    throw new Error("Invalid folder: missing drive_id or driveId");
  }

  try {
    const client = await getGraphClient(siteSettings, account);

    const driveId = folder.drive_id || folder.driveId;
    const itemId = folder.item_id || folder.itemId;

    if (!driveId || !itemId) {
      throw new Error("Invalid folder: missing drive_id or item_id");
    }

    // Fetch folder details to get the full path
    const folderResponse = await client
      .api(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`
      )
      .select("name,parentReference,webUrl")
      .get();

    // Build the full path by traversing up the parent chain
    const pathParts = [folderResponse.name];
    let currentItem = folderResponse;
    let maxDepth = 20; // Safety limit to prevent infinite loops

    // Traverse up the parent chain to build the full path
    while (currentItem.parentReference?.id && maxDepth > 0) {
      maxDepth--;

      // Check if we've reached root
      if (
        currentItem.parentReference.id === "root" ||
        currentItem.parentReference.driveId !== driveId
      ) {
        break;
      }

      // Get parent item
      const parentResponse = await client
        .api(
          `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(currentItem.parentReference.id)}`
        )
        .select("name,parentReference")
        .get();

      // If parent is root or has no parent, stop
      if (
        !parentResponse.parentReference?.id ||
        parentResponse.parentReference.id === "root"
      ) {
        break;
      }

      // Add parent item name to path parts
      pathParts.unshift(parentResponse.name);
      currentItem = parentResponse;
    }

    const folderPath = pathParts.join("/");

    // Fetch folder children
    const response = await client
      .api(
        `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children`
      )
      .get();

    const files = response?.value || [];

    return {
      path: folderPath,
      files,
    };
  } catch (error) {
    surfaceError(error, "cs_discourse_onedrive.refresh_error");
    throw error;
  }
}

/**
 * Persist folder link to the backend by calling our backend route
 * @param {number} topicId - Topic ID
 * @param {Object} folder - Folder object to save
 * @returns {Promise<null>} Promise resolving to null (folder saved)
 */
export async function persistFolder(topicId, folder) {
  const url = `/cs-discourse-onedrive/topics/${topicId}/folder`;

  try {
    await ajax(url, {
      type: "PUT",
      data: { folder },
    });
  } catch (error) {
    surfaceError(error, "cs_discourse_onedrive.save_error");
    throw error;
  }
}

/**
 * Remove folder link from the backend
 * @param {number} topicId - Topic ID
 * @returns {Promise<null>} Promise resolving to null (folder removed)
 */
export async function removeFolder(topicId) {
  const url = `/cs-discourse-onedrive/topics/${topicId}/folder`;

  try {
    await ajax(url, {
      type: "DELETE",
    });
  } catch (error) {
    surfaceError(error, "cs_discourse_onedrive.save_error");
    throw error;
  }
}
