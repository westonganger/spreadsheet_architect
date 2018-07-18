class CustomColumnsMethodPost < ActiveRecord::Base
  self.table_name = :posts

  include SpreadsheetArchitect

  def my_custom_spreadsheet_columns
    [
      :name,
      ['The Content', :content],
      :created_at, 
      ['Created At Date', created_at.to_date],
      [:asd, 'tadaaa']
    ]
  end

  SPREADSHEET_OPTIONS = {
    headers: true,
    spreadsheet_columns: :my_custom_spreadsheet_columns,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: 'My Project Export',
    column_styles: [],
    range_styles: [],
    merges: [],
    borders: [],
    column_types: []
  }

end
