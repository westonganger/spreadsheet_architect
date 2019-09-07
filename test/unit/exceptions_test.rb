require 'test_helper'

class ExceptionsTest < ActiveSupport::TestCase

  test "ArgumentError" do
    error = SpreadsheetArchitect::Exceptions::ArgumentError

    assert_raise error do
      conditional_row_styles = [{}]
      SpreadsheetArchitect::Utils::XLSX.conditional_styles_for_row(conditional_row_styles, true, true)
    end

    assert_raise error do
      conditional_row_styles = [{if: true, unless: true, styles: {}}]
      SpreadsheetArchitect::Utils::XLSX.conditional_styles_for_row(conditional_row_styles, true, true)
    end

    assert_raise error do
      conditional_row_styles = [{if: true, styles: false}]
      SpreadsheetArchitect::Utils::XLSX.conditional_styles_for_row(conditional_row_styles, true, true)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils.get_options({freeze: "A1", freeze_headers: true}, SpreadsheetArchitect)
    end

    assert_nothing_raised do
      SpreadsheetArchitect.default_options = {freeze_headers: true}
      SpreadsheetArchitect::Utils.get_options({freeze: "A1"}, SpreadsheetArchitect)
    end
  end

  test "NoDataError" do
    error = SpreadsheetArchitect::Exceptions::NoDataError

    assert error.new

    assert_raise error do
      SpreadsheetArchitect.to_csv(foo: :bar)
    end
  end

  test "MultipleDataSourcesError" do
    error = SpreadsheetArchitect::Exceptions::MultipleDataSourcesError

    assert error.new

    assert_raise error do
      SpreadsheetArchitect.to_csv(data: [], instances: [])
    end
  end

  test "OptionTypeError" do
    error = SpreadsheetArchitect::Exceptions::OptionTypeError

    assert error.new(:foobar_option).message
    assert_raise error do
      SpreadsheetArchitect.to_csv(data: {})
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(instances: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(headers: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(header_style: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(row_style: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(column_styles: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(range_styles: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(merges: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(borders: :foo)
    end
    assert_raise error do
      SpreadsheetArchitect.to_csv(column_widths: :foo)
    end
  end

  test "SpreadsheetColumnsNotDefinedError" do
    error = SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError

    assert_raise ArgumentError do
      error.new
    end

    assert error.new(SpreadsheetArchitect)

    class QuickBadClass
      include SpreadsheetArchitect
    end

    assert_raise error do
      QuickBadClass.to_csv(instances: [])
    end
  end

  test "InvalidColumnError" do
    error = SpreadsheetArchitect::Exceptions::InvalidColumnError

    assert_raise ArgumentError do
      error.new
    end

    assert error.new(:foobar_column)

    assert_raise error do
      SpreadsheetArchitect.to_xlsx(data: [[1]], column_styles: [{columns: 999}])
    end

    assert_raise error do
      SpreadsheetArchitect.to_xlsx(data: [[1]], column_styles: [{columns: 'ZZZ'}])
    end

    assert_raise error do
      SpreadsheetArchitect.to_xlsx(data: [[1]], column_styles: [{columns: [999]}])
    end

    assert_raise error do
      SpreadsheetArchitect.to_xlsx(data: [[1]], column_styles: [{columns: ('ZZX'..'ZZZ')}])
    end
  end

  test "InvalidRangeStylesOptionError" do
    error = SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError

    assert_raise ArgumentError do
      error.new
    end

    assert error.new(:columns, :bar_opt)
    assert error.new(:row, :bar_opt)
    assert error.new(:foo, :bar_opt)

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.range_hash_to_str({columns: :foo}, 1, 1)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.range_hash_to_str({rows: :foo}, 1, 1)
    end
  end

  test "InvalidRangeError" do
    error = SpreadsheetArchitect::Exceptions::InvalidRangeError

    assert_raise ArgumentError do
      error.new
    end

    errors = []
    errors.push error.new(:foo_type, :bar_opt)
    errors.push error.new(:columns, :bar_opt)
    errors.push error.new(:rows, :bar_opt)
    errors.push error.new(:format, :bar_opt)
    errors.push error.new(:type, :bar_opt)
    errors.push error.new(:foo_type, :bar_opt)

    assert_equal errors.count, errors.uniq.count
    assert error.new(:foo_type, :bar_opt)

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.verify_range(:foo, 1)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.verify_range("foo", 1)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.verify_range("foo:foo", 1)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.verify_range("A1:A2", 1)
    end

    assert_raise error do
      SpreadsheetArchitect::Utils::XLSX.verify_range("@1:A2", 1)
    end
  end

end
