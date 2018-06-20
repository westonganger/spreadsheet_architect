require 'test_helper'

class CsvTest < ActiveSupport::TestCase

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
      column_types: []
    }
  end

  def test_empty_model
    Post.delete_all
    File.open(Rails.root.join('tmp/empty_model.csv'),'w+b') do |f|
      f.write Post.to_csv
    end
  end

  def test_empty_sa
    File.open(Rails.root.join('tmp/empty_sa.csv'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_csv(data: [])
    end
  end

  def test_sa
    File.open(Rails.root.join('tmp/sa.csv'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_csv(headers: @headers, data: @data)
    end
  end

  def test_model
    File.open(Rails.root.join('tmp/model.csv'),'w+b') do |f|
      f.write Post.to_csv
    end
  end

  def test_options
    File.open(Rails.root.join('tmp/options.csv'),'w+b') do |f|
      f.write Post.to_csv(@options)
    end
  end

  def teardown
  end

end
