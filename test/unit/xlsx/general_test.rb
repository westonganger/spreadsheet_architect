require 'test_helper'

class XlsxGeneralTest < ActiveSupport::TestCase

  describe "header" do
    test "header: false" do
      spreadsheet = SpreadsheetArchitect.to_xlsx(headers: false, data: [[1]])
      
      save_file("xlsx/headers_false_test.xlsx", spreadsheet)
    end
  end

  test "use_zero_based_row_index" do
    spreadsheet = SpreadsheetArchitect.to_xlsx(
      use_zero_based_row_index: true,
      headers: [1, 2, 3],
      data: [
        [4,5,6],
        [7,8,9],
        [10,11,12],
      ],
      range_styles: [
        {range: {rows: 0, columns: 0}, styles: {bold: true}},
        {range: {rows: (2..3), columns: 0}, styles: {italic: true}},
      ],
      merges: [
        {range: {rows: 0, columns: 1..2}},
        {range: {rows: (2..3), columns: 2}},
      ],
      borders: [
        {range: {rows: 0, columns: :all}},
        {range: {rows: 3, columns: 0}},
      ],
    )

    ### TODO Verify contents
    
    save_file("xlsx/use_zero_based_row_index_test.xlsx", spreadsheet)
  end

  test "multisheet" do
    test_data = [[1,2,3], [4,5,6], [7,8,9]]

    package = Post.to_axlsx_package
    package = CustomPost.to_axlsx_package({sheet_name: 'Latest Projects'}, package)
    package = SpreadsheetArchitect.to_axlsx_package({data: test_data, sheet_name: 'Another Sheet'}, package)

    save_file("xlsx/multi_sheet.xlsx", package)
  end

  test "kitchen sink" do
    options = {
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Posted on','Posted At','Earnings']
      ],
      data: 50.times.map{|i| [i, "foobar-#{i}", 5.4*i, true, Date.today, Time.now]},
      header_style: {background_color: "000000", color: "FFFFFF", align: :center, font_size: 12, bold: true},
      row_style: {background_color: nil, color: "000000", align: :left, font_size: 12},
      sheet_name: 'Kitchen Sink',
      freeze_headers: true,

      column_styles: [
        {columns: 0, styles: {bold: true}},
        {columns: (1..3), styles: {color: "444444"}},
        {columns: [4], include_header: false, styles: {italic: true}}
      ],

      range_styles: [
        {range: "B2:C4", styles: {background_color: "CCCCCC"}},
        {range: {rows: 1, columns: :all}, styles: {bold: true}},
        {range: {rows: (20..25), columns: (1..4)}, styles: {italic: true}},
        {range: {rows: :all, columns: (3..4)}, styles: {color: "999999"}}
      ],

      conditional_row_styles: [
        {
          if: Proc.new{|row_data, row_index|
            row_index == 0 || row_data[0].to_i == 2
          }, 
          styles: {font_size: 18}
        },
        {
          unless: Proc.new{|row_data, row_index|
            row_index <= 10 || row_data[0].to_i == 15
          }, 
          styles: {align: :right}
        },
      ],

      borders: [
        {range: "B2:C4"},
        {range: "D6:D7", border_styles: {style: :dashDot, color: "333333"}},
        {range: {rows: (2..11), columns: :all}, border_styles: {edges: [:top,:bottom]}},
        {range: {rows: 3, columns: 4}, border_styles: {edges: [:top,:bottom]}}
      ],

      merges: [
        {range: "A1:C1"},
        {range: {rows: 2, columns: :all}},
        {range: {rows: (3..4), columns: (0..3)}}
      ],

      column_types: [
        :integer,
        :string,
        :float,
        :boolean,
        :date,
        :time,
        nil,
      ],
    }

    # Using Array Data
    file_data = SpreadsheetArchitect.to_xlsx(options)

    save_file("xlsx/kitchen_sink.xlsx", file_data)
  end

end
