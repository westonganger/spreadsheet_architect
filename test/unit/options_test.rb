require 'test_helper'

class OptionsTest < ActiveSupport::TestCase

  test "header: false" do
    assert SpreadsheetArchitect.to_xlsx(headers: false, data: [[1]])
  end

end
