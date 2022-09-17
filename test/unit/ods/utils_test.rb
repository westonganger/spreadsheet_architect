require 'test_helper'

class OdsUtilsTest < ActiveSupport::TestCase
  
  describe "get_ods_cell_type" do
    describe "without passing the type argument" do
      test "it gets the correct cell type" do
        assert_equal :string, SpreadsheetArchitect::Utils.get_ods_cell_type("foobar")
        assert_equal :float, SpreadsheetArchitect::Utils.get_ods_cell_type(1)
      end
    end

    describe "when passing the type argument" do
      test "returns the given type" do
        assert_equal "foobar", SpreadsheetArchitect::Utils.get_ods_cell_type(1, "foobar")
      end

      describe "hyperlinks" do
        test "returns :string" do
          assert_equal :string, SpreadsheetArchitect::Utils.get_ods_cell_type("foobar", :hyperlink)
          assert_equal :string, SpreadsheetArchitect::Utils.get_ods_cell_type(1, :hyperlink)
        end
      end
    end
  end

end
