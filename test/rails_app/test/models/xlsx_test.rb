require 'test_helper'

class XlsxTest < ActiveSupport::TestCase

  def setup
    @headers = ['test1', 'test2','test3']
    @data = [
      ['row1'],
      ['row2 c1', 'row2 c2'],
      ['the','data']
    ]
  end

  def test_empty_model
    Post.delete_all
    File.open('tmp/empty_model.xlsx','w+b') do |f|
      f.write Post.to_xlsx
    end
  end

  def test_empty_sa
    File.open('tmp/empty_sa.xlsx','w+b') do |f|
      f.write SpreadsheetArchitect.to_xlsx(data: [])
    end
  end

  def test_sa
    File.open('tmp/sa.xlsx','w+b') do |f|
      f.write SpreadsheetArchitect.to_xlsx(headers: @headers, data: @data)
    end
  end

  def test_model
    File.open('tmp/model.xlsx','w+b') do |f|
      f.write Post.to_xlsx
    end
  end
  
  def test_extreme
    headers = [
      ['Latest Posts'],
      ['Title','Category','Author','Posted on','Earnings']
    ]

    data = [
      50.times.map{|x| ['Title','Category','Author','Posted on','Earnings']}
    ]

    header_style = {}

    row_style = {}

    column_styles = [
      {columns: 0, styles: {bold: true}},
      {columns: (1..3), styles: {format_code: "$#,##0.00"}},
      {columns: [4,9], include_header: true, styles: {italic: true}}
    ]

    range_styles = [
      {range: "B2:C4", styles: {background_color: "CCCCCC"}}
    ]

    borders = [
      {range: "B2:C4"},
      {range: "D6:D7", border_styles: {style: :dashdot, color: "333333"}},
      {rows: (2..11), border_styles: {edges: [:top,:bottom]}},
      {rows: [1,3,5], columns: ('B'..'F')},
      {rows: 1},
      {columns: 0, border_styles: {edges: [:right], style: :thick}},
      {columns: (1..2)},
      {columns: ('D'..'F')},
      {columns: [4,6,8]},
      {columns: ['A','C','E']}

    ]

    merges = [
      {range: "A1:C1"}
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

    File.open('tmp/extreme.xlsx','w+b') do |f|
      f.write file_data
    end
  end

  def teardown
    DatabaseCleaner.clean
  end

end
