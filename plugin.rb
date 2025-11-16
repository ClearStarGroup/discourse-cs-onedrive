# frozen_string_literal: true

# name: cs-discourse-onedrive
# about: Clear Star Discourse Plugin to integrate OneDrive folders with topics
# meta_topic_id: TODO
# version: 0.0.1
# authors: Clear Star
# url: TODO
# required_version: 2.7.0

enabled_site_setting :cs_discourse_onedrive_enabled

register_asset "stylesheets/common/cs-discourse-onedrive.scss"
register_svg_icon "cloud"
register_svg_icon "file-pdf"
register_svg_icon "file-word"
register_svg_icon "file-excel"
register_svg_icon "file-powerpoint"
register_svg_icon "file-csv"
register_svg_icon "file-zipper"

module ::CsDiscourseOneDriveModule
  PLUGIN_NAME = "cs-discourse-onedrive"

  FOLDER_FIELD = "cs_discourse_onedrive_folder"

  def self.folder_from(topic)
    raw = topic.custom_fields[FOLDER_FIELD]
    return if raw.blank?

    raw.is_a?(String) ? JSON.parse(raw) : raw
  rescue JSON::ParserError
    nil
  end

  def self.serialize_for(topic, guardian)
    {
      folder: folder_from(topic),
      can_manage: guardian.can_edit?(topic&.first_post),
    }
  end
end

require_relative "lib/cs_discourse_onedrive_module/engine"

after_initialize do
  require_relative "app/controllers/cs_discourse_onedrive_module/auth_controller"
  require_relative "app/controllers/cs_discourse_onedrive_module/folders_controller"

  register_topic_custom_field_type(CsDiscourseOneDriveModule::FOLDER_FIELD, :json)
  TopicList.preloaded_custom_fields << CsDiscourseOneDriveModule::FOLDER_FIELD

  TopicView.on_preload do |topic_view|
    topic_view.topic&.custom_fields&.[](CsDiscourseOneDriveModule::FOLDER_FIELD)
  end

  add_model_callback Topic, :before_destroy do
    custom_fields.delete(CsDiscourseOneDriveModule::FOLDER_FIELD) if custom_fields.present?
  end

  add_to_serializer(:topic_view, :cs_discourse_onedrive) do
    CsDiscourseOneDriveModule.serialize_for(object.topic, scope)
  end

  add_to_serializer(
    :post,
    :cs_discourse_onedrive,
    include_condition: -> { object.is_first_post? }
  ) do
    CsDiscourseOneDriveModule.serialize_for(object.topic, scope)
  end
end
