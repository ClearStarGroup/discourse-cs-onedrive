# frozen_string_literal: true

module ::CsDiscourseOneDriveModule
  class ExamplesController < ::ApplicationController
    requires_plugin PLUGIN_NAME

    skip_before_action :preload_json, :check_xhr

    def index
      render html: "<h1>Hello, world!</h1>", layout: false
    end
  end
end
