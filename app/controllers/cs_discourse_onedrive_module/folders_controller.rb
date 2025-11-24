# frozen_string_literal: true

module ::CsDiscourseOneDriveModule
  class FoldersController < ::ApplicationController
    # Ensure the plugin is loaded before this controller is used
    requires_plugin PLUGIN_NAME

    # Disable Discourse layout rendering for this controller
    layout false

    # Skip preloading JSON; currently a no-op but could impact error pages
    skip_before_action :preload_json

    # Ensure the user is logged in
    before_action :ensure_logged_in

    # Find the topic
    before_action :find_topic
    def find_topic
      @topic = Topic.find_by(id: params[:topic_id])
      raise Discourse::NotFound unless @topic
    end

    # Ensure the user can manage the topic
    before_action :ensure_can_manage!
    def ensure_can_manage!
      guardian.ensure_can_edit!(@topic)
    end

    # ROUTE: update the folder for a topic
    def update
      # Parse params - require folder param, limit allowed fields, and convert to hash with symbol and string keys
      folder = params
        .require(:folder)
        .permit(:drive_id, :item_id, :name, :path, :web_url)
        .to_h
        .with_indifferent_access

      # Validate required fields
      if folder[:drive_id].blank? || folder[:item_id].blank?
        raise Discourse::InvalidParameters.new(
          I18n.t("cs_discourse_onedrive.errors.missing_folder_params"),
        )
      end

      # Check if folder already exists (to distinguish "set" from "changed")
      existing_folder = @topic.custom_fields[FOLDER_FIELD]
      is_change = existing_folder.present?

      # Extract folder name and URL
      folder_name = folder[:name]
      folder_url = folder[:web_url]
      if folder_name.blank? || folder_url.blank?
        raise Discourse::InvalidParameters.new(
          I18n.t("cs_discourse_onedrive.errors.missing_folder_params"),
        )
      end

      # Update topic custom field with parsed folder data
      @topic.custom_fields[FOLDER_FIELD] = folder
      @topic.save_custom_fields(true)

      # Create small action post (use add_moderator_post to preserve custom_fields)
      action_code = is_change ? "onedrive_folder_changed" : "onedrive_folder_linked"
      @topic.add_moderator_post(
        current_user,
        nil,
        post_type: Post.types[:small_action],
        action_code: action_code,
        bump: false,
        silent: true,
        custom_fields: {
          "onedrive_folder_name" => folder_name.to_s,
          "onedrive_folder_path" => folder_url.to_s,
        },
      )

      # Return success response
      head :ok
    end

    # ROUTE: delete the folder for a topic
    def delete
      # Capture folder info before deletion using the helper method for consistency
      existing_folder = CsDiscourseOneDriveModule.folder_from(@topic)
      folder_name = existing_folder&.dig("name") || existing_folder&.dig(:name)
      folder_url = existing_folder&.dig("web_url") || existing_folder&.dig(:web_url)

      # Delete folder custom field from topic
      @topic.custom_fields.delete(FOLDER_FIELD)
      @topic.save_custom_fields(true)

      # Create small action post (only if folder existed and has a name)
      if folder_name.present?
        @topic.add_moderator_post(
          current_user,
          nil,
          post_type: Post.types[:small_action],
          action_code: "onedrive_folder_cleared",
          bump: false,
          silent: true,
          custom_fields: {
            "onedrive_folder_name" => folder_name.to_s,
            "onedrive_folder_path" => folder_url.to_s,
          },
        )
      end

      # Return success response
      head :ok
    end
  
  end
end

