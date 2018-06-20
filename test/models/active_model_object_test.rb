require "test_helper"

class ActiveModelObjectTest < ActiveSupport::TestCase

  def setup
    @path = Rails.root.join('tmp/active_model_object')
    FileUtils.mkdir_p(@path)
  end

  def test_csv
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      ActiveModelObject.to_csv
    end

    instances = 20.times.map{|_| ActiveModelObject.new(name: :asd, content: :asd, created_at: Time.now)}
    data = ActiveModelObject.to_csv(instances: instances)

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write data
    end
  end

  def test_ods
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      ActiveModelObject.to_ods
    end

    instances = 20.times.map{|_| ActiveModelObject.new(name: :asd, content: :asd, created_at: Time.now)}
    data = ActiveModelObject.to_ods(instances: instances)

    File.open(File.join(@path, 'ods.ods'), 'w+b') do |f|
      f.write data
    end
  end

  def test_xlsx
    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      ActiveModelObject.to_xlsx
    end

    instances = 20.times.map{|_| ActiveModelObject.new(name: :asd, content: :asd, created_at: Time.now)}
    data = ActiveModelObject.to_xlsx(instances: instances)

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write data
    end
  end

  def teardown
  end
end
