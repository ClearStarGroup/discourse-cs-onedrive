import { eq } from "discourse/truth-helpers";
import CsDiscourseOnedrivePanel from "discourse/plugins/discourse-cs-onedrive/discourse/components/cs-discourse-onedrive-panel";

const OneDriveSection = <template>
  {{yield @outletArgs}}
  {{#if (eq @outletArgs.post.post_number 1)}}
    <CsDiscourseOnedrivePanel @post={{@outletArgs.post}} />
  {{/if}}
</template>;

export default OneDriveSection;
