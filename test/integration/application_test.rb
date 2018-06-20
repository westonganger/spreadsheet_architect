### FOR RAILS TESTS

require 'test_helper'

class ApplicationTest < ActionDispatch::IntegrationTest
  fixtures :all

  def before
    @path = Rails.root.join('tmp/controller_tests')
    FileUtils.mkdir_p(@path)
  end

  test "Renders html action" do
    get '/spreadsheet/xlsx'
    assert_response :success
    #assert @response.body.include?("<h1>Hello World!</h1>")
    assert_match("<h1>Hello World!</h1>", @response.body)
  end

  test "Renders xlsx to string" do
    xlsx_str = ApplicationController.new.render_to_string("/spreadsheet/xlsx.xlsx", locals: {})

    assert_not_nil xlsx_str
    assert_match("<h1>Hello World!</h1>", @response.body)
  end

  test "Renders sample xlsx action" do
    get '/spreadsheet/xlsx', params: {format: :xlsx}
    assert_response :success
    assert_not @response.body.include?("<h1>Hello World!</h1>")

    #reader = PDF::Reader.new(StringIO.new(@response.body))
    assert_nil(reader)

    #assert_equal(1, reader.page_count)

    #page_str = reader.pages[0].to_s

    #assert page_str.include?('Hello World!')
    #assert_not page_str.include?("<h1>Hello World!</h1>")
  end

  test "" do
    assert_not_equal 'smoke', 'test'
  end

  def test_csv
    get '/spreadsheet/csv'
    assert_response :success
    File.open(File.join(@path, 'tmp/controller_tests/csv.csv'), 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_ods
    get '/spreadsheet/ods'
    assert_response :success

    File.open(File.join(@path, 'tmp/controller_tests/ods.ods'), 'w+b') do |f|
      f.write @response.body
    end
  end

  def test_xlsx
    get '/spreadsheet/xlsx'
    assert_response :success
    File.open(File.join(@path, 'tmp/controller_tests/xlsx.xlsx'), 'w+b') do |f|
      f.write @response.body
    end

    get '/spreadsheet/alt_xlsx', params: {format: :xlsx}
    assert_response :success
    File.open(File.join(@path, 'tmp/controller_tests/alt_xlsx.xlsx'), 'w+b') do |f|
      f.write @response.body
    end
  end
end
