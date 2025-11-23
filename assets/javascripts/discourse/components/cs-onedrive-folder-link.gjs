import Component from "@glimmer/component";

export default class CsOnedriveFolderLink extends Component {
  <template>
    <p>
      {{#if this.args.folder.web_url}}
        <a
          href={{this.args.folder.web_url}}
          target="_blank"
          rel="noopener"
        >
          {{#if this.args.folderPath}}
            {{this.args.folderPath}}
          {{else}}
            {{this.args.folder.name}}
          {{/if}}
        </a>
      {{else}}
        {{#if this.args.folderPath}}
          {{this.args.folderPath}}
        {{else}}
          {{this.args.folder.name}}
        {{/if}}
      {{/if}}
    </p>
  </template>
}

