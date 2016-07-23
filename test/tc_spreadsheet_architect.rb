#!/usr/bin/env ruby -w

lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)

require 'minitest'

require 'spreadsheet_architect'

require 'minitest/autorun'

class TestSpreadsheetArchitect < MiniTest::Test
  def setup
    require_relative 'helper'

    Minitest::Assertions.module_eval do
      alias_method :eql, :assert_equal
    end

    SpreadsheetArchitect.default_options = {
      headers: true,
      header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      sheet_name: 'My Project Export',
      column_styles: [],
      range_styles: [],
      borders: []
    }
  end

  def test_utils_methods
    eql(SpreadsheetArchitect::Utils.str_humanize('my_project_export'), 'My Project Export')
    eql(SpreadsheetArchitect::Utils.str_humanize('My Project Export'), 'My Project Export')
    eql(SpreadsheetArchitect::Utils.str_humanize('TBS report'), 'TBS Report')

    eql(SpreadsheetArchitect::Utils.get_type('string'), :string)
    eql(SpreadsheetArchitect::Utils.get_type(123.01), :float)
    eql(SpreadsheetArchitect::Utils.get_type(BigDecimal('123.01')), :float)
    eql(SpreadsheetArchitect::Utils.get_type(10), :integer)
    eql(SpreadsheetArchitect::Utils.get_type(:test), :symbol)

    eql(SpreadsheetArchitect::Utils.get_type(Time.now.to_date), :date)
    eql(SpreadsheetArchitect::Utils.get_type(DateTime.now), :time)
    eql(SpreadsheetArchitect::Utils.get_type(Time.now), :time)

    #SpreadsheetArchitect::Utils.get_cell_data(options, klass)

    #SpreadsheetArchitect::Utils.get_options(options, klass)
      
    eql(SpreadsheetArchitect::Utils.convert_styles_to_axlsx({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true}), {bg_color: '333333', fg_color: '000000', alignment: {horizontal: true}, b: true, sz: 14, i: true, u: true, test: true})
  end

  def test_csv_methods
    headers = ['test1', 'test2','test3']
    data = [
      ['row1'],
      ['row2 c1', 'row2 c2'],
      ['the','data']
    ]

    #Post.to_csv
    SpreadsheetArchitect.to_csv(headers: headers, data: data)

    #Post.to_ods
    SpreadsheetArchitect.to_ods(headers: headers, data: data)

    #Post.to_rodf_spreadsheet
    SpreadsheetArchitect.to_rodf_spreadsheet(headers: headers, data: data)

    #Post.to_xlsx
    SpreadsheetArchitect.to_xlsx(headers: headers, data: data)

    #Post.to_axlsx_package({data: data, headers: headers})
    package = SpreadsheetArchitect.to_axlsx_package({data: data, headers: headers})
    SpreadsheetArchitect.to_axlsx_package({data: data, headers: headers, sheet_name: 'The Sheet Name'}, package)
  end
end
