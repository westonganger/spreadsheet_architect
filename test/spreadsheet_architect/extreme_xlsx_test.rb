require "test_helper"

class ExtremeXlsxTest < ActiveSupport::TestCase

  def setup
  end

  def teardown
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

    File.open(File.join('tmp/extreme.xlsx'),'w+b') do |f|
      f.write file_data
    end
  end

end
