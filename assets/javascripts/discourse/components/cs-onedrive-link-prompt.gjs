import Component from "@glimmer/component";
import DButton from "discourse/components/d-button";
import { i18n } from "discourse-i18n";

export default class CsOnedriveLinkPrompt extends Component {
  <template>
    {{#if this.args.canManage}}
      <DButton
        @label="discourse_cs_onedrive.link_folder"
        @action={{this.args.onLinkFolder}}
        @disabled={{this.args.linking}}
        class="btn btn-primary"
      />
    {{else}}
      <p>{{i18n "discourse_cs_onedrive.no_folder_linked"}}</p>
    {{/if}}
  </template>
}
