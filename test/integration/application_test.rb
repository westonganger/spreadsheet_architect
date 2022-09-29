### FOR RAILS TESTS

require 'test_helper'

class ApplicationTest < ActionDispatch::IntegrationTest
  def setup
    @path = TMP_PATH.join("integration")
    FileUtils.mkdir_p(@path)
  end

  def test_csv
    get '/spreadsheets/csv', params: {format: :csv}
    assert_response :success

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_ods
    get '/spreadsheets/ods'
    assert_response :success

    File.open(File.join(@path, 'ods.ods'), 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_xlsx
    get '/spreadsheets/xlsx'
    assert_response :success

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_alt_xlsx
    get '/spreadsheets/alt_xlsx', params: {format: :xlsx}
    assert_response :success

    File.open(File.join(@path, 'alt_xlsx.xlsx'), 'w+b') do |f|
      f.write @response.body
    end
  end
end
