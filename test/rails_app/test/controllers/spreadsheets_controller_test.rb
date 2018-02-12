require_relative "../test_helper"

class SpreadsheetsControllerTest < ActionController::TestCase

  def setup
    DatabaseCleaner.start

    FileUtils.mkdir_p('tmp/controller_tests')
  end

  def test_csv
    get :csv
    assert_response :success
    File.open('tmp/controller_tests/csv.csv', 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_ods
    get :ods
    assert_response :success

    File.open('tmp/controller_tests/ods.ods', 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_xlsx
    get :xlsx
    assert_response :success
    File.open('tmp/controller_tests/xlsx.xlsx', 'w+b') do |f|
      f.write @response.body
    end

    get :alt_xlsx, params: {format: :xlsx}
    assert_response :success
    File.open('tmp/controller_tests/alt_xlsx.xlsx', 'w+b') do |f|
      f.write @response.body
    end
  end

  def teardown
    DatabaseCleaner.clean
  end

end
