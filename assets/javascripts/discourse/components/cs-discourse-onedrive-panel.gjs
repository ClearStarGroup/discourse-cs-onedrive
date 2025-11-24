import Component from "@glimmer/component";
import { action } from "@ember/object";
import { tracked } from "@glimmer/tracking";
import { service } from "@ember/service";
import loadingSpinner from "discourse/helpers/loading-spinner";
import icon from "discourse/helpers/d-icon";
import { i18n } from "discourse-i18n";
import {
  acquireAccessToken,
  getAccounts,
  handleRedirectResult,
  signOut as authSignOut,
  surfaceError,
} from "discourse/plugins/discourse-cs-onedrive/discourse/lib/onedrive-auth-service";
import { openPicker } from "discourse/plugins/discourse-cs-onedrive/discourse/lib/onedrive-picker-service";
import {
  loadFiles,
  persistFolder,
  removeFolder,
} from "discourse/plugins/discourse-cs-onedrive/discourse/lib/onedrive-api-service";
import CsOnedriveFileList from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-file-list";
import CsOnedriveFolderHeader from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-folder-header";
import CsOnedriveFolderLink from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-folder-link";
import CsOnedriveSignInPrompt from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-sign-in-prompt";
import CsOnedriveLinkPrompt from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-link-prompt";
import CsOnedriveErrorAlert from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-onedrive-error-alert";

export default class CsDiscourseOnedrivePanel extends Component {
  @service currentUser;
  @service siteSettings;

  // The last error message that occurred.  If set, we show an error message.
  @tracked errorMessage = null;

  // The OneDrive account that is signed in.  If not set, we only show a sign in prompt.
  @tracked account = null;

  // The OneDrive folder that is linked to the topic.  If not set, we only show a link folder prompt.
  @tracked folder = null;

  // Whether we're loading data from OneDrive.  If true, we only show a loading spinner.
  @tracked loading = false;

  // The loaded data
  @tracked folderPath = null;
  @tracked files = null;

  // Whether we're linking a new folder.  If true, we disable the link folder buttons.
  @tracked linking = false;

  constructor() {
    super(...arguments);
    this._initialize();
  }

  // Is the plugin enabled? If not, we'll do nothing and show nothing.
  get isEnabled() {
    return this.siteSettings.discourse_cs_onedrive_enabled;
  }

  // Is all the plugin configuration present and valid? If not we'll just show an error.
  get isConfigured() {
    return (
      this.siteSettings.discourse_cs_onedrive_client_id &&
      this.siteSettings.discourse_cs_onedrive_client_id.trim().length > 0 &&
      this.siteSettings.discourse_cs_onedrive_tenant_id &&
      this.siteSettings.discourse_cs_onedrive_tenant_id.trim().length > 0 &&
      this.siteSettings.discourse_cs_onedrive_sharepoint_base_url &&
      this.siteSettings.discourse_cs_onedrive_sharepoint_base_url.trim()
        .length > 0 &&
      this.siteSettings.discourse_cs_onedrive_sharepoint_site_name &&
      this.siteSettings.discourse_cs_onedrive_sharepoint_site_name.trim()
        .length > 0
    );
  }

  // Topic id of the current topic
  get topicId() {
    return this.args.post?.topic_id || this.args.post?.topicId;
  }

  // Does the current user have permissions to set/change the linked folder?
  get canManage() {
    return this.args.post?.can_edit;
  }

  async _initialize() {
    // Do nothing if the plugin is not enabled or not configured
    if (!this.isEnabled || !this.isConfigured) {
      return;
    }

    // Handle any pending redirect results (clears "interaction in progress" state)
    await handleRedirectResult(this.siteSettings);

    // Get accounts from MSAL cache (populated by callback page after redirect)
    const accounts = await getAccounts(this.siteSettings);
    if (!this.account && accounts.length > 0) {
      this.account = accounts[0];
    }

    // Initialize folder from args (synchronous, happens before async operations)
    this.folder =
      this.args.post?.discourse_cs_onedrive?.folder ||
      this.args.post?.topic?.discourse_cs_onedrive?.folder ||
      null;

    // Load data if we have a linked folder and a signed in account
    if (this.folder && this.account) {
      await this._loadData();
    }
  }

  async _loadData() {
    if (!this.folder) {
      return;
    }

    this.errorMessage = null;
    this.loading = true;

    try {
      const result = await loadFiles(
        this.folder,
        this.siteSettings,
        this.account
      );

      this.files = result.files;
      this.folderPath = result.path;
    } catch (error) {
      const errorMessage = surfaceError(
        error,
        "discourse_cs_onedrive.refresh_error"
      );
      this.errorMessage = errorMessage;
      throw error;
    } finally {
      this.loading = false;
    }
  }

  @action
  async signIn() {
    try {
      await acquireAccessToken(this.siteSettings, this.account, true);
      if (this.folder) {
        await this._loadData();
      }
    } catch (error) {
      this.errorMessage = surfaceError(
        error,
        "discourse_cs_onedrive.auth_error"
      );
    }
  }

  @action
  async signOut() {
    if (!this.account) {
      return;
    }

    try {
      await authSignOut(this.siteSettings, this.account);
    } catch (error) {
      this.errorMessage = surfaceError(
        error,
        "discourse_cs_onedrive.auth_error"
      );
    }

    this.account = null;
    this.files = null;
    this.folderPath = null;
  }

  @action
  async refreshFiles() {
    await this._loadData();
  }

  @action
  async linkFolder() {
    if (!this.canManage) {
      return;
    }

    this.linking = true;
    this.errorMessage = null;

    try {
      const selection = await openPicker(this.siteSettings, this.account);
      await persistFolder(this.topicId, selection);
      this.folder = selection;
      await this._loadData();
    } catch (error) {
      if (error?.message !== "cancelled") {
        this.errorMessage = i18n("discourse_cs_onedrive.picker_error");
      }
    } finally {
      this.linking = false;
    }
  }

  @action
  async removeFolder() {
    if (!this.canManage) {
      return;
    }

    if (!confirm(i18n("discourse_cs_onedrive.confirm_remove"))) {
      return;
    }

    this.errorMessage = null;

    try {
      await removeFolder(this.topicId);
      this.folder = null;
      this.files = null;
      this.folderPath = null;
    } catch (error) {
      this.errorMessage = surfaceError(
        error,
        "discourse_cs_onedrive.remove_error"
      );
    }
  }

  <template>
    {{#if this.isEnabled}}
      <div class="post__row row">
        <div class="topic-avatar cs-onedrive-avatar">
          {{icon "cloud"}}
        </div>

        <div class="post__body topic-body clearfix">
          <CsOnedriveFolderHeader
            @folder={{this.folder}}
            @account={{this.account}}
            @loading={{this.loading}}
            @linking={{this.linking}}
            @onRefresh={{this.refreshFiles}}
            @onChangeFolder={{this.linkFolder}}
            @onRemoveFolder={{this.removeFolder}}
            @onSignOut={{this.signOut}}
          />

          <div class="post__regular regular post__contents contents">
            <div class="cooked">
              {{#unless this.isConfigured}}
                <CsOnedriveErrorAlert
                  @errorMessage={{i18n "discourse_cs_onedrive.not_configured"}}
                />
              {{else if this.errorMessage}}
                <CsOnedriveErrorAlert @errorMessage={{this.errorMessage}} />
              {{else unless this.account}}
                <CsOnedriveSignInPrompt @onSignIn={{this.signIn}} />
              {{else unless this.folder}}
                <CsOnedriveLinkPrompt
                  @canManage={{this.canManage}}
                  @linking={{this.linking}}
                  @onLinkFolder={{this.linkFolder}}
                />
              {{else if this.loading}}
                <div class="cs-onedrive-panel__spinner">
                  {{loadingSpinner}}
                </div>
              {{else}}
                <CsOnedriveFolderLink
                  @folder={{this.folder}}
                  @folderPath={{this.folderPath}}
                />
                <CsOnedriveFileList @files={{this.files}} />
              {{/unless}}
            </div>
          </div>
        </div>
      </div>
    {{/if}}
  </template>
}
