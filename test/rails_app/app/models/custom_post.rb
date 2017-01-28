class CustomPost < ActiveRecord::Base
  self.table_name = :posts

  include SpreadsheetArchitect

  def spreadsheet_columns
    [
      :name,
      ['The Content', :content],
      :created_at, 
      ['Created At 2', created_at.strftime("%Y-%m-%d")],
      [:asd, 'tadaaa']
    ]
  end
end
