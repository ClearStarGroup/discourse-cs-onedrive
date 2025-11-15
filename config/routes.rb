# frozen_string_literal: true

CsDiscourseOneDriveModule::Engine.routes.draw do
  get "/examples" => "examples#index"
  # define routes here
end

Discourse::Application.routes.draw { mount ::CsDiscourseOneDriveModule::Engine, at: "cs-discourse-onedrive" }
