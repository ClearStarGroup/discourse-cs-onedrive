import { eq } from "discourse/truth-helpers";

const OneDriveSection = <template>
  {{yield @outletArgs}}
  {{#if (eq @outletArgs.post.post_number 1)}}
    <div class="onedrive-section">
      Hello
    </div>
  {{/if}}
</template>;

export default OneDriveSection;

