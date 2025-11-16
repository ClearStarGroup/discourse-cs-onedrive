# frozen_string_literal: true

module ::CsDiscourseOneDriveModule
  class FoldersController < ::ApplicationController
    requires_plugin PLUGIN_NAME

    layout false
    skip_before_action :preload_json, :check_xhr
    before_action :ensure_logged_in
    before_action :find_topic
    before_action :ensure_can_manage!

    def update
      folder = permitted_folder

      if folder[:drive_id].blank? || folder[:item_id].blank?
        raise Discourse::InvalidParameters.new(
          I18n.t("cs_discourse_onedrive.errors.missing_folder_params"),
        )
      end

      folder[:linked_at] = Time.zone.now.iso8601

      @topic.custom_fields[FOLDER_FIELD] = folder
      @topic.save_custom_fields(true)

      render json: CsDiscourseOneDriveModule.serialize_for(@topic, guardian), layout: false
    end

    def destroy
      @topic.custom_fields.delete(FOLDER_FIELD)
      @topic.save_custom_fields(true)

      render json: CsDiscourseOneDriveModule.serialize_for(@topic, guardian), layout: false
    end

    private

    def permitted_folder
      params
        .require(:folder)
        .permit(:drive_id, :item_id, :name, :path, :web_url)
        .to_h
        .with_indifferent_access
    end

    def ensure_can_manage!
      guardian.ensure_can_edit!(@topic.first_post)
    end

    def guardian
      @guardian ||= Guardian.new(current_user)
    end

    def find_topic
      @topic = Topic.find(params[:topic_id])
    end
  end
end

