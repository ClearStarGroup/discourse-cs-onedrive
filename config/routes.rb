# frozen_string_literal: true

CsDiscourseOneDriveModule::Engine.routes.draw do
  get "/auth/callback" => "auth#callback"
  put "/topics/:topic_id/folder" => "folders#update"
  delete "/topics/:topic_id/folder" => "folders#delete"
end

Discourse::Application.routes.draw { mount ::CsDiscourseOneDriveModule::Engine, at: "cs-discourse-onedrive" }
