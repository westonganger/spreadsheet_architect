require 'test_helper'

class ExceptionsTest < ActiveSupport::TestCase

  test "NoDataError" do
    assert SpreadsheetArchitect::Exceptions::NoDataError.new
  end

  test "MultipleDataSourcesError" do
    assert SpreadsheetArchitect::Exceptions::MultipleDataSourcesError.new
  end

  test "IncorrectTypeError" do
    assert SpreadsheetArchitect::Exceptions::IncorrectTypeError.new
    assert SpreadsheetArchitect::Exceptions::IncorrectTypeError.new(:foobar_option)
  end

  test "SpreadsheetColumnsNotDefinedError" do
    assert SpreadsheetArchitect::Exceptions::IncorrectTypeError.new
    assert SpreadsheetArchitect::Exceptions::IncorrectTypeError.new(SpreadsheetArchitect)
  end

  test "InvalidColumnError" do
    assert_raise do
      SpreadsheetArchitect::Exceptions::InvalidColumnError.new
    end

    assert SpreadsheetArchitect::Exceptions::InvalidColumnError.new(:foobar_column)
  end

  test "InvalidRangeStylesOptionError" do
    assert_raise do
      SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError.new
    end

    assert SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError.new(:foo_type, :bar_opt)
  end

  test "BadRangeError" do
    assert_raise do
      SpreadsheetArchitect::Exceptions::BadRangeError.new
    end

    errors = []
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:foo_type, :bar_opt)
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:columns, :bar_opt)
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:rows, :bar_opt)
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:format, :bar_opt)
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:type, :bar_opt)
    errors.push SpreadsheetArchitect::Exceptions::BadRangeError.new(:foo_type, :bar_opt)

    assert_equal errors.count, errors.uniq.count
    assert SpreadsheetArchitect::Exceptions::BadRangeError.new(:foo_type, :bar_opt)
  end

end
