# frozen_string_literal: true

CsDiscourseOneDriveModule::Engine.routes.draw do
  put "/topics/:topic_id/folder" => "folders#update"
  delete "/topics/:topic_id/folder" => "folders#destroy"
  get "/auth/callback" => "auth#callback"
  get "/auth/picker_callback" => "auth#picker_callback"
end

Discourse::Application.routes.draw { mount ::CsDiscourseOneDriveModule::Engine, at: "cs-discourse-onedrive" }
