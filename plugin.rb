# frozen_string_literal: true

# name: discourse-cs-onedrive
# about: Clear Star Discourse Plugin to integrate OneDrive folders with topics
# meta_topic_id: TODO
# version: 0.0.1
# authors: Clear Star
# url: TODO
# required_version: 2.7.0

# Which of our settings turns the plugin on or off
enabled_site_setting :discourse_cs_onedrive_enabled

# Register custom SCSS for the plugin
register_asset "stylesheets/common/discourse-cs-onedrive.scss"

# Register SVG icons for the plugin
register_svg_icon "cloud"
register_svg_icon "file-pdf"
register_svg_icon "file-word"
register_svg_icon "file-excel"
register_svg_icon "file-powerpoint"
register_svg_icon "file-csv"
register_svg_icon "file-zipper"

module ::DiscourseCsOnedriveModule
  # The name of the plugin
  PLUGIN_NAME = "discourse-cs-onedrive"

  # The name of the custom field we use to store the folder for a topic
  FOLDER_FIELD = "onedrive_folder"

  # Helper method to get the folder for a topic (called only by serialize_for below)
  def self.folder_from(topic)
    raw = topic.custom_fields[FOLDER_FIELD]
    return if raw.blank?

    raw.is_a?(String) ? JSON.parse(raw) : raw
  rescue JSON::ParserError
    nil
  end

  # Helper method to serialize the folder for a topic (called by the topic_view serializers below)
  def self.serialize_for(topic)
    {
      folder: folder_from(topic),
    }
  end
end

# Load the engine to enable routes to be registered (happens before after_initialize called)
require_relative "lib/engine"

after_initialize do
  # Register the custom field type we use to store the folder for a topic
  register_topic_custom_field_type(DiscourseCsOnedriveModule::FOLDER_FIELD, :json)

  # Ensure this field is included in topic objects to avoid N+1 queries
  TopicList.preloaded_custom_fields << DiscourseCsOnedriveModule::FOLDER_FIELD

  # Delete the folder for a topic when it is destroyed
  add_model_callback Topic, :before_destroy do
    custom_fields.delete(DiscourseCsOnedriveModule::FOLDER_FIELD) if custom_fields.present?
  end

  # Add the linked folder to the topic object where serializing topics
  add_to_serializer(:topic_view, :discourse_cs_onedrive) do
    DiscourseCsOnedriveModule.serialize_for(object.topic)
  end

  # Add our custom fields to the post custom fields allowlist
  # This ensures they're included when serializing posts in topic views
  TopicView.add_post_custom_fields_allowlister do |user, topic|
    ["onedrive_folder_name", "onedrive_folder_path"]
  end

  # Add custom fields for OneDrive folder name and path to post serializer
  # These are used by small action posts to display folder links
  add_to_serializer(
    :post,
    :onedrive_folder_name,
    include_condition: -> { post_custom_fields["onedrive_folder_name"].present? },
  ) { post_custom_fields["onedrive_folder_name"] }

  add_to_serializer(
    :post,
    :onedrive_folder_path,
    include_condition: -> { post_custom_fields["onedrive_folder_path"].present? },
  ) { post_custom_fields["onedrive_folder_path"] }
end
