## Draft Overview

**Requirement**

- Insert a OneDrive section after the main post of a topic, before replies.
- Require every viewer to authenticate with their own OneDrive account (via MSAL in the browser) before interacting with the panel.
- When a OneDrive folder is already linked to the topic, display its files (with refresh/change/delete actions) using the current viewer’s Microsoft access token.
- When no folder is linked, prompt an authorised viewer to select/link a OneDrive folder and persist the association with the topic.
- Respect per-topic scope (other topics remain unaffected).

**High-Level Solution**

- Extend the Discourse topic view with a plugin outlet rendering the OneDrive panel.
- Store the OneDrive folder metadata against the topic (topic custom fields or associated table) and expose the linkage through the topic/post serializers.
- Each viewer completes Microsoft OAuth via MSAL; the plugin calls Microsoft Graph directly from the browser with the viewer’s access tokens (no server-side token storage or proxy).
- Update the panel reactively when authentication completes, a folder is selected, the viewer clicks refresh, or the viewer lacks access to the folder.

**Technical Notes**

- UI: Ember component plugged into `post:after` outlet (or new outlet); distinct states for “unauthenticated” (MSAL sign-in prompt), “no folder linked” (picker CTA), “folder linked with access” (file list + refresh/change/delete actions), and “folder linked but viewer lacks access” (error messaging). Styling should rely on core Discourse classes/variables to match the default themes out of the box.
- Server: Store OneDrive metadata in topic custom fields (single folder per topic). Expose values via topic/post serializers and provide endpoints to save/update/delete the custom-field payload; no proxy endpoints for Microsoft Graph.
- Auth: Import `@azure/msal-browser` via the wrapper in `discourse/lib/msal-browser`, so the module is bundled with the plugin (no CDN dependency) and MSAL APIs are available synchronously.
- File listing: Import `@microsoft/microsoft-graph-client` via `discourse/lib/microsoft-graph-client`, then call `Client.init` directly inside the component before issuing Graph requests.
- Folder selection: Integrate the official OneDrive/Graph JavaScript picker (`microsoft-file-picker`) to browse and select folders; ensure callback persists drive + item IDs alongside display name/path; the folder owner still needs to grant Microsoft-level access to intended viewers.
- Permissions: reuse Discourse’s built-in ability_to_edit_post checks—only users who can edit the topic’s first post (topic owner, staff, or others with edit rights) can link/change/delete folders. All other viewers see the linked folder (if any) but no management controls; evaluation of extra group gating is unnecessary.
- Background jobs: not required. All Graph calls happen client-side with the viewer’s tokens; rely on user-triggered refresh or per-viewer caching. Webhook-driven nudges can be considered later without server-side jobs.

**Open Questions**

None

**Updating MSAL/Graph**

1. Bump the versions of `@azure/msal-browser` / `@microsoft/microsoft-graph-client` in `package.json`, then run `pnpm install`.
2. Copy the browser bundles from `node_modules` into the plugin’s vendor directory:
   - `cp node_modules/@azure/msal-browser/lib/msal-browser.min.js assets/javascripts/vendor/msal-browser.min.js`
   - `cp node_modules/@microsoft/microsoft-graph-client/lib/graph-js-sdk.js assets/javascripts/vendor/microsoft-graph-client.js`
3. Restart the Ember CLI dev server (or rebuild assets) so the new vendor files are fingerprinted.
4. Smoke-test the panel (sign-in, folder selection, file listing) to confirm nothing regressed.

**Useful Docs**

- Install Discourse for Development with Docker - https://meta.discourse.org/t/install-discourse-for-development-using-docker/102009
- Plugin development - https://meta.discourse.org/t/developing-discourse-plugins-part-1-create-a-basic-plugin/30515

**Key Discourse Docker Commands**

- init: d/boot_dev --init
- stop: d/shutdown_dev
- start: d/boot_dev

- start rails: d/rails s
- start ember: d/ember-cli

http://localhost:4200

**TODO**

- Remove pnpm packages, and update spec and instructions to download from unpkg
- Remove picker callback?
- Move picker code to stand-alone lib
- Remove dependancy on graph lib (create our own lib)
- Rationalise msal to a single lib (rather than two)
- Sort out styling and layout (table, better buttons, use Discourse styles)
- Constrain picker (maybe to site configured in settings)
- Review of all code
