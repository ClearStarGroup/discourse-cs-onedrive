import Component from "@glimmer/component";
import DButton from "discourse/components/d-button";
import { i18n } from "discourse-i18n";

export default class CsOnedriveSignInPrompt extends Component {
  <template>
    <DButton
      @label="cs_discourse_onedrive.sign_in"
      @action={{this.args.onSignIn}}
      class="btn btn-primary"
    />
  </template>
}

