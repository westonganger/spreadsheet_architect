require 'test_helper'

class GeneralTest < ActiveSupport::TestCase

  test "Version accessible by default" do
    assert_not_nil SpreadsheetArchitect::VERSION
  end

  test "Constants dont change" do
    x = SpreadsheetArchitect.default_options.to_s
    SpreadsheetArchitect.to_xlsx(headers: [[1]], data: [[1]], header_style: {b: false}, row_style: {background_color: '000000'})
    assert_equal(x, SpreadsheetArchitect.default_options.to_s)
  end

end
