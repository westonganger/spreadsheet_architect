require "test_helper"

class SpreadsheetArchitectTest < ActiveSupport::TestCase

  def setup
    @path = Rails.root.join('tmp/spreadsheet_architect')
    FileUtils.mkdir_p(@path)

    @headers = ['test1', 'test2','test3']
    @data = [
      ['row1','test1'],
      ['row2 c1', nil, '123'],
      ['the',Date.today,Time.now]
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
      column_types: [:string, :date, :date]
    }
  end

  def test_csv
    File.open(File.join(@path, 'csv.csv'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_csv(headers: @headers, data: @data)
    end
  end

  def test_empty_csv
    File.open(File.join(@path, 'empty.csv'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_csv(data: [])
    end
  end

  def test_empty_ods
    File.open(File.join(@path, 'empty.ods'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_ods(data: [])
    end
  end

  def test_ods
    File.open(File.join(@path, 'ods.ods'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_ods(headers: @headers, data: @data)
    end
  end

  def test_empty_xlsx
    File.open(File.join(@path, 'empty.xlsx'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_xlsx(data: [])
    end
  end

  def test_xlsx
    File.open(File.join(@path, 'xlsx.xlsx'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_xlsx(headers: @headers, data: @data)
    end
  end

  def test_extreme_xlsx
    headers = [
      ['Latest Posts'],
      ['Title','Category','Author','Posted on','Earnings']
    ]

    data = 50.times.map{|i| [i, i+i, i+i+i, i+i+i+i, i+i+i+i+i]}

    header_style = {}

    row_style = {}

    column_styles = [
      {columns: 0, styles: {bold: true}},
      {columns: (1..3), styles: {format_code: "$#,##0.00"}},
      {columns: [4], include_header: true, styles: {italic: true}}
    ]

    range_styles = [
      {range: "B2:C4", styles: {background_color: "CCCCCC"}},
      {range: {rows: 1, columns: :all}, styles: {bold: true}},
      {range: {rows: (0..5), columns: (1..4)}, styles: {italic: true}},
      {range: {rows: :all, columns: (3..4)}, styles: {color: "999999"}}
    ]

    borders = [
      {range: "B2:C4"},
      {range: "D6:D7", border_styles: {style: :dashDot, color: "333333"}},
      {range: {rows: (2..11), columns: :all}, border_styles: {edges: [:top,:bottom]}},
      {range: {rows: 3, columns: 4}, border_styles: {edges: [:top,:bottom]}}
    ]

    merges = [
      {range: "A1:C1"},
      {range: {rows: 0, columns: :all}},
      {range: {rows: (0..1), columns: (0..3)}}
    ]


    # Using Array Data
    file_data = SpreadsheetArchitect.to_xlsx({
      headers: headers,
      data: data,
      header_style: header_style,
      row_style: row_style,
      column_styles: column_styles,
      range_styles: range_styles,
      borders: borders,
      merges: merges
    })

    File.open(File.join(@path, 'extreme.xlsx'),'w+b') do |f|
      f.write file_data
    end
  end

  def teardown
  end
end
