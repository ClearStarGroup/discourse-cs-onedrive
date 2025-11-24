import { htmlSafe } from "@ember/template";
import { autoUpdatingRelativeAge } from "discourse/lib/formatter";
import { withPluginApi } from "discourse/lib/plugin-api";
import { i18n } from "discourse-i18n";

const ONEDRIVE_ACTION_CODES = [
  "onedrive_folder_linked",
  "onedrive_folder_changed",
  "onedrive_folder_cleared",
];

export default {
  name: "extend-small-action-onedrive",
  initialize() {
    withPluginApi((api) => {
      // Add icon for OneDrive action codes
      api.addPostSmallActionIcon("onedrive_folder_linked", "cloud");
      api.addPostSmallActionIcon("onedrive_folder_changed", "cloud");
      api.addPostSmallActionIcon("onedrive_folder_cleared", "cloud");

      // Extend the small action component to add custom placeholders for folder name and URL
      api.modifyClass(
        "component:post/small-action",
        (Superclass) =>
          class extends Superclass {
            get description() {
              // Only apply custom logic for OneDrive action codes
              if (ONEDRIVE_ACTION_CODES.includes(this.code)) {
                const when = this.createdAt
                  ? autoUpdatingRelativeAge(this.createdAt, {
                      format: "medium-with-ago-and-on",
                    })
                  : "";

                // Get folder name and URL from serialized attributes
                const folderName = this.args.post.onedrive_folder_name || "";
                const folderUrl = this.args.post.onedrive_folder_path || "";

                // Pass custom placeholders to i18n
                return htmlSafe(
                  i18n(`action_codes.${this.code}`, {
                    who: "",
                    when,
                    folder_name: folderName,
                    folder_url: folderUrl,
                  })
                );
              }

              // For all other action codes, use default behavior
              return super.description;
            }
          },
        { pluginId: "discourse-cs-onedrive" }
      );
    });
  },
};
