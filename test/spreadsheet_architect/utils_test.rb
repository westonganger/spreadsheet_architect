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

  def test_constants_dont_change
    x = SpreadsheetArchitect.default_options.to_s
    SpreadsheetArchitect.to_xlsx(headers: [[1]], data: [[1]], header_style: {b: false}, row_style: {background_color: '000000'})
    assert_equal(x, SpreadsheetArchitect.default_options.to_s)
  end

  def test_str_humanize
    assert_equal(SpreadsheetArchitect::Utils.str_humanize('my_project_export'), 'My Project Export')
    assert_equal(SpreadsheetArchitect::Utils.str_humanize('My Project Export'), 'My Project Export')
    assert_equal(SpreadsheetArchitect::Utils.str_humanize('TBS report'), 'TBS Report')
  end

  def test_get_type
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type('string'), :string)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(123.01), :float)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(BigDecimal('123.01')), :float)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(10), :integer)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(:test), :string)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(Time.now.to_date), :date)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(DateTime.now), :time)
    assert_equal(SpreadsheetArchitect::Utils::XLSX.get_type(Time.now), :time)
  end

  def test_get_cell_data
    SpreadsheetArchitect::Utils.get_cell_data(@options, SpreadsheetArchitect)

    SpreadsheetArchitect::Utils.get_cell_data(@options, Post)
  end

  def test_get_options
    SpreadsheetArchitect::Utils.get_options(@options, SpreadsheetArchitect)

    SpreadsheetArchitect::Utils.get_options(@options, Post)
  end
      
  def test_convert_styles_to_axlsx
    xlsx_styles = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true})
    assert_equal(xlsx_styles, {bg_color: '333333', fg_color: '000000', alignment: {horizontal: true}, b: true, sz: 14, i: true, u: true, test: true})
  end
  
  def test_convert_styles_to_ods
    ods_styles = SpreadsheetArchitect::Utils.convert_styles_to_ods({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true})
    assert_equal(ods_styles, {'cell' => {'background-color' => '#333333'}, 'text' => {'color' => '#000000', 'align' => true, 'font-weight' => 'bold', 'font-size' => 14, 'font-style' => 'italic', 'text-underline-type' => 'single', 'text-underline-style' => 'solid'}})
  end

  def teardown
  end
end
