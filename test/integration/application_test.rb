### FOR RAILS TESTS

require 'test_helper'

class ApplicationTest < ActionDispatch::IntegrationTest
  def setup
    @path = TMP_PATH.join("integration")
    FileUtils.mkdir_p(@path)
  end

  def test_csv
    get '/spreadsheets/csv'
    assert_response :success

    save_file('integration/test.csv', @response.body)
  end

  def test_ods
    get '/spreadsheets/ods'
    assert_response :success

    save_file('integration/test.ods', @response.body)
  end

  def test_xlsx
    get '/spreadsheets/xlsx'
    assert_response :success

    save_file('integration/test.xlsx', @response.body)
  end

  def test_respond_with
    get '/spreadsheets/test_respond_with', params: {format: :xlsx}
    assert_response :success

    save_file('integration/test_respond_with.xlsx', @response.body)
  end
end
