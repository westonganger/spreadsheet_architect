require 'test_helper'

class UtilsTest < ActiveSupport::TestCase
  klass = SpreadsheetArchitect::Utils

  def setup
    @options = {
      header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      sheet_name: 'My Project Export',
      column_styles: [],
      range_styles: [],
      merges: [],
      borders: [],
      column_types: [],

      headers: ['test1', 'test2','test3'],
      data: [
        ['row1'],
        ['row2 c1', 'row2 c2'],
        ['the','data']
      ]
    }
  end

  def teardown
  end

  test "get_cell_data" do
    klass.get_cell_data(@options, SpreadsheetArchitect)

    klass.get_cell_data(@options, Post)
  end

  test "get_options" do
    klass.get_options(@options, SpreadsheetArchitect)

    klass.get_options(@options, Post)
  end
  
  test "convert_styles_to_ods" do
    ods_styles = klass.convert_styles_to_ods({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true})
    assert_equal(ods_styles, {'cell' => {'background-color' => '#333333'}, 'text' => {'color' => '#000000', 'align' => true, 'font-weight' => 'bold', 'font-size' => 14, 'font-style' => 'italic', 'text-underline-type' => 'single', 'text-underline-style' => 'solid'}})
  end

  test "deep_clone" do
    skip("TODO")
  end

  test "is_ar_model" do
    skip("TODO")
  end

  test "str_humanize" do
    assert_equal(klass.str_humanize('my_project_export'), 'My Project Export')
    assert_equal(klass.str_humanize('My Project Export'), 'My Project Export')
    assert_equal(klass.str_humanize('TBS report'), 'TBS Report')
  end

  test "check_type" do
    skip("TODO")
  end

  test "check_options_types" do
    skip("TODO")
  end

  test "stringify_keys" do
    skip("TODO")
  end
end
