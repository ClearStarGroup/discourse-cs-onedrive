import Component from "@glimmer/component";
import DButton from "discourse/components/d-button";
import icon from "discourse/helpers/d-icon";
import { i18n } from "discourse-i18n";

export default class CsOnedriveFolderHeader extends Component {
  <template>
    <div class="topic-meta-data">
      <div class="names">
        <span class="first username cs-onedrive-header">
          {{i18n "cs_discourse_onedrive.section_header"}}
        </span>
      </div>

      {{#if this.args.account}}
        <div class="post-infos">
          <div class="actions">
            {{#if this.args.folder}}
              <DButton
                @icon="arrows-rotate"
                @title="cs_discourse_onedrive.refresh"
                @action={{this.args.onRefresh}}
                @disabled={{this.args.loading}}
                class="btn no-text btn-icon post-action-menu__refresh btn-flat"
              />
              <DButton
                @icon="folder"
                @title="cs_discourse_onedrive.change_folder"
                @action={{this.args.onChangeFolder}}
                @disabled={{this.args.linking}}
                class="btn no-text btn-icon post-action-menu__change-folder btn-flat"
              />
              <DButton
                @icon="trash-can"
                @title="cs_discourse_onedrive.remove_folder"
                @action={{this.args.onRemoveFolder}}
                class="btn no-text btn-icon post-action-menu__remove btn-flat btn-danger"
              />
            {{/if}}
            <DButton
              @icon="right-from-bracket"
              @title="cs_discourse_onedrive.sign_out"
              @action={{this.args.onSignOut}}
              class="btn no-text btn-icon post-action-menu__logout btn-flat"
            />
          </div>
        </div>
      {{/if}}
    </div>
  </template>
}
