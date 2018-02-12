require "test/test_helper"

class CustomPostTest < ActiveSupport::TestCase

  def setup
    @path = File.join(Rails.root, 'tmp/custom_posts')
    FileUtils.mkdir_p(@path)

    DatabaseCleaner.start
  end

  def test_csv
    data = CustomPost.to_csv

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write data
    end
  end

  def test_ods
    data = CustomPost.to_ods

    File.open(File.join(@path, 'ods.ods'), 'w+b') do |f|
      f.write data
    end
  end

  def test_xlsx
    data = CustomPost.all.to_xlsx

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write data
    end
  end

  def test_empty
    Post.destroy_all
    
    File.open(File.join(@path, 'empty.xlsx'), 'w+b') do |f|
      f.write Post.to_xlsx
    end
  end

  def test_constants_dont_change
    x = CustomPost::SPREADSHEET_OPTIONS.to_s
    SpreadsheetArchitect.to_xlsx(headers: [[1]], data: [[1]], header_style: {b: false}, row_style: {background_color: '000000'})
    assert_equal(x, CustomPost::SPREADSHEET_OPTIONS.to_s)
  end


  def teardown
    DatabaseCleaner.clean
  end
end
