# frozen_string_literal: true

# name: cs-discourse-onedrive
# about: Clear Star Discourse Plugin to integrate OneDrive folders with topics
# meta_topic_id: TODO
# version: 0.0.1
# authors: Clear Star
# url: TODO
# required_version: 2.7.0

enabled_site_setting :cs_discourse_onedrive_enabled

module ::CsDiscourseOneDriveModule
  PLUGIN_NAME = "cs-discourse-onedrive"
end

require_relative "lib/cs_discourse_onedrive_module/engine"

after_initialize do
  require_relative "app/controllers/cs_discourse_onedrive_module/examples_controller"
end
