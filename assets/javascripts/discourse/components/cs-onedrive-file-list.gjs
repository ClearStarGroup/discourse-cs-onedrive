import Component from "@glimmer/component";
import icon from "discourse/helpers/d-icon";
import formatDate from "discourse/helpers/format-date";
import { i18n } from "discourse-i18n";
import {
  getFileTypeIcon,
  getFileTypeName,
  formatFileSize,
} from "discourse/plugins/cs-discourse-onedrive/discourse/lib/file-type-utils";

export default class CsOnedriveFileList extends Component {
  get files() {
    return this.args.files || [];
  }

  get hasFiles() {
    return this.args.files && this.args.files.length > 0;
  }

  <template>
    {{#if this.hasFiles}}
      <table class="cs-onedrive-panel__table">
        <thead>
          <tr>
            <th></th>
            <th>{{i18n "cs_discourse_onedrive.file_name"}}</th>
            <th>{{i18n "cs_discourse_onedrive.last_modified"}}</th>
            <th>{{i18n "cs_discourse_onedrive.file_type"}}</th>
            <th>{{i18n "cs_discourse_onedrive.file_size"}}</th>
          </tr>
        </thead>
        <tbody>
          {{#each this.files as |file|}}
            <tr>
              <td class="cs-onedrive-panel__icon-cell">
                {{icon (getFileTypeIcon file)}}
              </td>
              <td class="cs-onedrive-panel__name-cell">
                {{#if file.webUrl}}
                  <a href={{file.webUrl}} target="_blank" rel="noopener">
                    {{file.name}}
                  </a>
                {{else}}
                  {{file.name}}
                {{/if}}
              </td>
              <td class="cs-onedrive-panel__date-cell">
                {{#if file.lastModifiedDateTime}}
                  {{formatDate file.lastModifiedDateTime format="medium"}}
                {{else}}
                  —
                {{/if}}
              </td>
              <td class="cs-onedrive-panel__type-cell">
                {{getFileTypeName file}}
              </td>
              <td class="cs-onedrive-panel__size-cell">
                {{formatFileSize file.size}}
              </td>
            </tr>
          {{/each}}
        </tbody>
      </table>
    {{else if this.args.files}}
      <p>{{i18n "cs_discourse_onedrive.no_files"}}</p>
    {{/if}}
  </template>
}
