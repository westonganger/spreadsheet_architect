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

  SPREADSHEET_OPTIONS = {
    headers: true,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: 'My Project Export',
    column_styles: [],
    range_styles: [],
    merges: [],
    borders: [],
    column_types: []
  }
end
