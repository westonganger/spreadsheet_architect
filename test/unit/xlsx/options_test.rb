require 'test_helper'

class XlsxOptionsTest < ActiveSupport::TestCase

  test "header: false" do
    spreadsheet = SpreadsheetArchitect.to_xlsx(headers: false, data: [[1]])
    
    File.open(TMP_PATH.join("headers_false_test.xlsx"), 'w+b') do |f|
      f.write(spreadsheet)
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
    
    File.open(TMP_PATH.join("use_zero_based_row_index_test.xlsx"), 'w+b') do |f|
      f.write(spreadsheet)
    end
  end

end
