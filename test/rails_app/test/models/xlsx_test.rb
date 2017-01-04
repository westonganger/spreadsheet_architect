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

  def teardown
    DatabaseCleaner.clean
  end

end
