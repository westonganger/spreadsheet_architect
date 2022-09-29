class SpreadsheetsController < ApplicationController

  respond_to :csv, :xlsx, :ods
  
  def csv
    render csv: Post.to_csv
  end

  def ods
    render ods: Post.to_ods
  end

  def xlsx
    render xlsx: Post.to_xlsx, filename: 'Posts'
  end

  def test_respond_with
    if Rails::VERSION::MAJOR >= 5
      @posts = Post.all

      respond_with @posts
    end
  end
  
  def test_xlsx
    headers = [
      ['Latest Posts'],
      ['Title','Category','Author','Posted on','Earnings','Title','Category','Author','Posted on','Earnings']
    ]

    data = 50.times.map{|x| ['Title','Category','Author','Posted on','Earnings','Title','Category','Author','Posted on','Earnings'] }

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

    render xlsx: SpreadsheetArchitect.to_xlsx({
      headers: headers,
      data: data,
      header_style: header_style,
      row_style: row_style,
      column_styles: column_styles,
      range_styles: range_styles,
      borders: borders,
      merges: merges
    })
  end

end
