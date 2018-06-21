require "test_helper"

class CustomPostTest < ActiveSupport::TestCase

  def setup
    @path = Rails.root.join('tmp/custom_posts')
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
    data = CustomPost.to_csv

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write data
    end
  end

  def test_empty_ods
    File.open(File.join(@path, 'empty.ods'),'w+b') do |f|
      f.write CustomPost.where(id: 0).to_ods
    end
  end

  def test_ods
    File.open(File.join(@path, 'ods.ods'),'w+b') do |f|
      f.write CustomPost.to_ods
    end
  end

  def test_ods_options
    File.open(File.join(@path, 'options.ods'),'w+b') do |f|
      f.write CustomPost.to_ods(@options)
    end
  end

  def test_xlsx
    data = CustomPost.all.to_xlsx

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write data
    end
  end

  def test_empty
    File.open(File.join(@path, 'empty.xlsx'), 'w+b') do |f|
      f.write Post.where(id: 0).to_xlsx
    end
  end

  def test_constants_dont_change
    x = CustomPost::SPREADSHEET_OPTIONS.to_s
    SpreadsheetArchitect.to_xlsx(headers: [[1]], data: [[1]], header_style: {b: false}, row_style: {background_color: '000000'})
    assert_equal(x, CustomPost::SPREADSHEET_OPTIONS.to_s)
  end


  def teardown
  end
end
