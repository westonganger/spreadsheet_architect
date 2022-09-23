require 'test_helper'

class OdsMultiSheetTest < ActiveSupport::TestCase

  def setup
    @test_data = [[1,2,3], [4,5,6], [7,8,9]]
  end

  test "ods" do
    spreadsheet = Post.to_rodf_spreadsheet
    spreadsheet = CustomPost.to_rodf_spreadsheet({sheet_name: 'Latest Projects'}, spreadsheet)
    spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({data: @test_data, sheet_name: 'Another Sheet'}, spreadsheet)

    File.open(VERSIONED_BASE_PATH.join("multi_sheet.ods"),'w+b') do |f|
      f.write spreadsheet.bytes
    end
  end

end
