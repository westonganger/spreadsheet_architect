require 'rails/generators'

module SpreadsheetArchitect
  class AddDefaultOptionsGenerator < Rails::Generators::Base

    def add_default_options
      create_file "config/initializers/spreadsheet_architect.rb", <<eos
SpreadsheetArchitect.default_options = {
  headers: true,
  header_style: {background_color: "AAAAAA", color: "FFFFFF", align: :center, bold: false, font_name: 'Arial', font_size: 10, italic: false, underline: false},
  row_style: {background_color: nil, color: "000000", align: :left, bold: false, font_name: 'Arial', font_size: 10, italic: false, underline: false},
  column_styles: [],
  range_styles: [],
  borders: []
}
eos
    end

  end
end
