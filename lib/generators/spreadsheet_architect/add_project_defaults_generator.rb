require 'rails/generators'

module SpreadsheetArchitect
  class AddProjectDefaultsGenerator < Rails::Generators::Base

    def add_project_defaults
      create_file "config/initializers/spreadsheet_architect.rb", "SpreadsheetArchitect.module_eval do\n  const_set('SPREADSHEET_OPTIONS', {\n    headers: true,\n    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},\n    row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},\n    sheet_name: 'My Project Export'\n  })\nend"
    end

  end
end
