require "test/test_helper"

class BadPlainRubyObjectTest < ActiveSupport::TestCase

  def setup
    DatabaseCleaner.start
  end

  def test_csv
    assert_raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError do
      BadPlainRubyObject.to_csv(instances: [])
    end
  end

  def test_ods
    assert_raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError do
      BadPlainRubyObject.to_ods(instances: [])
    end
  end

  def test_xlsx
    assert_raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError do
      BadPlainRubyObject.to_xlsx(instances: [])
    end
  end

  def teardown
    DatabaseCleaner.clean
  end
end
