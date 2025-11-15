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
- Auth: Use `@azure/msal-browser` to manage authentication in the client; tokens live in browser storage/memory; handle silent token acquisition and fallbacks for interactive login.
- File listing: Microsoft Graph `/drives/{id}/items/{itemId}/children` called directly from the browser via `@microsoft/microsoft-graph-client`; implement client-side memoisation/buffering to reduce repeated calls while honouring permissions; bust cache when viewer clicks refresh.
- Folder selection: Integrate the official OneDrive/Graph JavaScript picker (`microsoft-file-picker`) to browse and select folders; ensure callback persists drive + item IDs alongside display name/path; the folder owner still needs to grant Microsoft-level access to intended viewers.
- Permissions: reuse Discourse’s built-in ability_to_edit_post checks—only users who can edit the topic’s first post (topic owner, staff, or others with edit rights) can link/change/delete folders. All other viewers see the linked folder (if any) but no management controls; evaluation of extra group gating is unnecessary.
- Background jobs: not required. All Graph calls happen client-side with the viewer’s tokens; rely on user-triggered refresh or per-viewer caching. Webhook-driven nudges can be considered later without server-side jobs.

**Open Questions**

None

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