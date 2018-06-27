require 'test_helper'

class UtilsTest < ActiveSupport::TestCase

  def setup
    @headers = ['test1', 'test2','test3']
    @data = [
      ['row1'],
      ['row2 c1', 'row2 c2'],
      ['the','data']
    ]
    @options = {
      headers: true,
      header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      sheet_name: 'My Project Export',
      column_styles: [],
      range_styles: [],
      merges: [],
      borders: [],
      column_types: [],
      data: []
    }
  end

  test "get_type" do
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type('string'), :string)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(123.01), :float)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(BigDecimal('123.01')), :float)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(10), :integer)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(:test), :string)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(Time.now.to_date), :date)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(DateTime.now), :time)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(Time.now), :time)
  end
      
  test "convert_styles_to_axlsx" do
    xlsx_styles = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true})
    assert_equal(xlsx_styles, {bg_color: '333333', fg_color: '000000', alignment: {horizontal: true}, b: true, sz: 14, i: true, u: true, test: true})
  end

  test "range_hash_to_str" do
    skip("TODO")
  end

  test "verify_range" do
    skip("TODO")
  end

  test "symbolize_keys" do
    skip("TODO")
  end

  def teardown
  end
end
