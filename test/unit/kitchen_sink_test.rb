require "test_helper"

class KitchenSinkTest < ActiveSupport::TestCase

  def setup
    @options = {
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Posted on','Earnings']
      ],
      data: 50.times.map{|i| [i, i*2, i*3, i*4, i*5, i*6, i*7, i*8]},
      header_style: {background_color: "000000", color: "FFFFFF", align: :center, font_size: 12, bold: true, italic: true},
      row_style: {background_color: nil, color: "000000", align: :left, font_size: 12},
      sheet_name: 'Kitchen Sink'
    }
  end

  def teardown
  end

  def test_xlsx
    @options.merge!({
      column_styles: [
        {columns: 0, styles: {bold: true}},
        {columns: (1..3), styles: {format_code: "$#,##0.00"}},
        {columns: [4], include_header: true, styles: {italic: true}}
      ],

      range_styles: [
        {range: "B2:C4", styles: {background_color: "CCCCCC"}},
        {range: {rows: 1, columns: :all}, styles: {bold: true}},
        {range: {rows: (0..5), columns: (1..4)}, styles: {italic: true}},
        {range: {rows: :all, columns: (3..4)}, styles: {color: "999999"}}
      ],

      conditional_row_styles: [
        {
          if: Proc.new{|row_data, row_index|
            row_index == 0 || row_data[0] == :foo
          }, 
          styles: {font_size: 14}
        },
        {
          unless: Proc.new{|row_data, row_index|
            row_index == 0 || row_data[0] == :foo
          }, 
          styles: {align: :right}
        },
      ],

      borders: [
        {range: "B2:C4"},
        {range: "D6:D7", border_styles: {style: :dashDot, color: "333333"}},
        {range: {rows: (2..11), columns: :all}, border_styles: {edges: [:top,:bottom]}},
        {range: {rows: 3, columns: 4}, border_styles: {edges: [:top,:bottom]}}
      ],

      merges: [
        {range: "A1:C1"},
        {range: {rows: 2, columns: :all}},
        {range: {rows: (3..4), columns: (0..3)}}
      ],

      column_types: [
        :string,
        :integer,
        :float,
        :boolean,
        nil,
      ],
    })

    # Using Array Data
    file_data = SpreadsheetArchitect.to_xlsx(@options)

    File.open(VERSIONED_BASE_PATH.join("kitchen_sink.xlsx"),'w+b') do |f|
      f.write file_data
    end
  end

  def test_ods
    @options.merge!({
      column_types: [
        :string,
        :float,
        :percent,
        :currency,
        :date,
        :time,
        nil
      ],
    })

    # Using Array Data
    file_data = SpreadsheetArchitect.to_ods(@options)

    File.open(VERSIONED_BASE_PATH.join("kitchen_sink.ods"),'w+b') do |f|
      f.write file_data
    end
  end

end
