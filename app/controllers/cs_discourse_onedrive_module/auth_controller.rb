# frozen_string_literal: true

module ::CsDiscourseOneDriveModule
  class AuthController < ::ApplicationController
    # Ensure the plugin is loaded before this controller is used
    requires_plugin PLUGIN_NAME 

    # Disable Discourse layout rendering for this controller
    layout false 

    # Skip preloading JSON; currently a no-op but could impact error pages
    skip_before_action :preload_json
    
    # Skip enforcement of XHR/JSON as callbacks are regular brower GET requests
    skip_before_action :check_xhr 

    # Skip CSRF verification as callback won't include our CSRF token
    skip_before_action :verify_authenticity_token 

    def callback
      # View will handle the MSAL redirect processing
      render template: "cs_discourse_onedrive_module/auth/callback", layout: false
    end
  end
end

