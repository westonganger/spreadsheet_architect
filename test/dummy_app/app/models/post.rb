class Post < ActiveRecord::Base
  include SpreadsheetArchitect

  attr_accessor :name, :content, :created_at

end
