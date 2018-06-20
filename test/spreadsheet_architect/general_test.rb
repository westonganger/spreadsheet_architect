require 'test_helper'

class GeneralTest < ActiveSupport::TestCase

  test "Version accessible by default" do
    assert_not_nil SpreadsheetArchitect::VERSION
  end

end
