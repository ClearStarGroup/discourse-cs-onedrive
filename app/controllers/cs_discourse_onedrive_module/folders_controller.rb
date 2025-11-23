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

      # Update topic custom field with parsed folder data
      @topic.custom_fields[FOLDER_FIELD] = folder
      @topic.save_custom_fields(true)

      # Return success response
      head :ok
    end

    # ROUTE: delete the folder for a topic
    def delete
      # Delete folder custom field from topic
      @topic.custom_fields.delete(FOLDER_FIELD)
      @topic.save_custom_fields(true)

      # Return success response
      head :ok
    end
  
  end
end

