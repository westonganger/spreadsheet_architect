class ActiveModelObject
  include ActiveModel::Model
  include SpreadsheetArchitect

  attr_accessor :name, :content, :created_at

  def spreadsheet_columns
    [
      :name,
      :content,
      :created_at
    ]
  end
end
