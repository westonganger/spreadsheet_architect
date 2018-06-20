require "test_helper"

class PlainRubyObjectTest < ActiveSupport::TestCase

  def setup
    @path = Rails.root.join('tmp/plain_ruby_object')
    FileUtils.mkdir_p(@path)
  end

  def test_csv
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      SpreadsheetArchitect.to_csv
    end

    instances = 3.times.map{|_| PlainRubyObject.new}
    data = SpreadsheetArchitect.to_csv(instances: instances)

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write data
    end
  end

  def test_ods
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      SpreadsheetArchitect.to_ods
    end

    instances = 3.times.map{|_| PlainRubyObject.new}
    data = SpreadsheetArchitect.to_ods(instances: instances)

    File.open(File.join(@path, 'ods.ods'), 'w+b') do |f|
      f.write data
    end
  end

  def test_xlsx
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      SpreadsheetArchitect.to_xlsx
    end

    instances = 3.times.map{|_| PlainRubyObject.new}
    data = SpreadsheetArchitect.to_xlsx(instances: instances)

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write data
    end
  end

  def teardown
  end
end
