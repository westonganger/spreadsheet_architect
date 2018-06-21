class LegacyPlainRubyObject
  include SpreadsheetArchitect ### old syntax, not in readme anymore, should easily remain compatible though

  attr_accessor :name, :content, :created_at

  def spreadsheet_columns
    [
      ['Name', :name],
      :content,
      ['Object ID', (object_id)]
    ]
  end

end
