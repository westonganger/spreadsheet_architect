require 'test_helper'

class SpreadsheetArchitect::Utils::OdsTest < ActiveSupport::TestCase
  klass = SpreadsheetArchitect::Utils::ODS
  
  describe "get_cell_type" do
    describe "without passing the type argument" do
      test "it gets the correct cell type" do
        assert_equal :string, klass.get_cell_type("foobar")
        assert_equal :float, klass.get_cell_type(1)
      end
    end

    describe "when passing the type argument" do
      test "returns the given type" do
        assert_equal "foobar", klass.get_cell_type(1, "foobar")
      end

      describe "hyperlinks" do
        test "returns :string" do
          assert_equal :string, klass.get_cell_type("foobar", :hyperlink)
          assert_equal :string, klass.get_cell_type(1, :hyperlink)
        end
      end
    end
  end

  describe "convert_styles" do
    test "convert_styles_to_ods" do
      ods_styles = klass.convert_styles({
        background_color: '333333', 
        color: '000000', 
        align: true, 
        bold: true, 
        font_size: 14, 
        italic: true, 
        underline: true, 
        test: true
      })

      assert_equal(ods_styles, {
        'cell' => {
          'background-color' => '#333333'
        }, 
        'text' => {
          'color' => '#000000', 
          'align' => true, 
          'font-weight' => 'bold', 
          'font-size' => 14, 
          'font-style' => 'italic', 
          'text-underline-type' => 'single', 
          'text-underline-style' => 'solid'
        }
      })
    end
  end

end
