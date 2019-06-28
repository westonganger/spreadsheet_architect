# rake install:local from root
require 'spreadsheet_architect'

base_spreadsheet_styles =
      {
        header_style: {
          background_color: "663496",
          color:            "FFFFFF",
          align:            :center,
          font_name:        'Arial',
          font_size:        20,
          bold:             true,
          italic:           false,
          underline:        false
        },
        row_style:    {
          background_color: nil,
          color:            "000000",
          align:            :left,
          font_name:        'Arial',
          font_size:        18,
          bold:             false,
          italic:           false,
          underline:        false,
          format_code:      nil
        },
        # column_styles:      [],
        # column_types:       [], # Valid types for XLSX are :string, :integer, :float, :boolean, nil = auto determine.
        # column_widths:      [], # use nil for auto-fil
      }


headers      = %w[one two three four five six seven eight nine ten]
test_data    = []

100.times do |row|
  row = []
  10.times do |col|
    row << col
  end
  test_data << row
end


options =
 {
    headers: headers,
    data: test_data,
  }.merge(base_spreadsheet_styles)

fixed_options = options.dup
fixed_options[:sheet_name] = 'Fixed'
fixed_options[:header_style][:fixed_top_left] = true

package = SpreadsheetArchitect.to_axlsx_package(fixed_options)


not_fixed_options = options.dup
not_fixed_options[:sheet_name] = 'Not Fixed'
not_fixed_options[:header_style][:fixed_top_left] = false

package = SpreadsheetArchitect.to_axlsx_package(not_fixed_options, package)


File.open("fixed_top_left.xlsx",'w+b') do |f|
  f.write package.to_stream.read
end
