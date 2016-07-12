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
    SpreadsheetArchitect::Utils.str_humanize

    #SpreadsheetArchitect::Utils.get_type(value, type=nil, last_run=false)

    SpreadsheetArchitect::Utils.get_cell_data(options, klass)

    SpreadsheetArchitect::Utils.get_options(options, klass)
      
    SpreadsheetArchitect::Utils.convert_styles_to_axlsx
  end

  def test_csv_methods
    #Post.to_csv
    SpreadsheetArchitect.to_csv

    #Post.to_ods
    SpreadsheetArchitect.to_ods

    #Post.to_rodf_spreadsheet
    SpreadsheetArchitect.to_rodf_spreadsheet

    #Post.to_xlsx
    SpreadsheetArchitect.to_xlsx

    #Post.to_axlsx('package', {data: data, headers: headers})
    SpreadsheetArchitect.to_axlsx('package', {data: data, headers: headers})

    #Post.to_axlsx('sheet', {data: data, headers: headers})
    SpreadsheetArchitect.to_axlsx('sheet', {data: data, headers: headers})
  end
end
