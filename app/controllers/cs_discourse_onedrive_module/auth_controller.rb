# frozen_string_literal: true

module ::CsDiscourseOneDriveModule
  class AuthController < ::ApplicationController
    requires_plugin PLUGIN_NAME
    layout false
    skip_before_action :preload_json, :check_xhr
    skip_before_action :verify_authenticity_token

    def callback
      # View will handle the MSAL redirect processing
      render template: "cs_discourse_onedrive_module/auth/callback", layout: false
    end

    def picker_callback
      # Simple callback page for OneDrive picker - the picker SDK handles everything
      render template: "cs_discourse_onedrive_module/auth/picker_callback", layout: false
    end
  end
end

