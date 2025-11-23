import Component from "@glimmer/component";

export default class CsOnedriveErrorAlert extends Component {
  <template>
    <div class="alert alert-error">
      {{this.args.errorMessage}}
    </div>
  </template>
}
