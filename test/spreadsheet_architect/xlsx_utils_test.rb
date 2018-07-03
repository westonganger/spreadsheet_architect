require 'test_helper'

class XlsxUtilsTest < ActiveSupport::TestCase
  klass = SpreadsheetArchitect::Utils::XLSX

  def setup
  end

  def teardown
  end

  test "get_type" do
    assert_equal(klass.get_type('string'), :string)
    assert_equal(klass.get_type(123.01), :float)
    assert_equal(klass.get_type(BigDecimal('123.01')), :float)
    assert_equal(klass.get_type(10), :integer)
    assert_equal(klass.get_type(:test), :string)
    assert_equal(klass.get_type(Time.now.to_date), :date)
    assert_equal(klass.get_type(DateTime.now), :time)
    assert_equal(klass.get_type(Time.now), :time)
  end
      
  test "convert_styles_to_axlsx" do
    xlsx_styles = klass.convert_styles_to_axlsx({background_color: '333333', color: '000000', align: true, bold: true, font_size: 14, italic: true, underline: true, test: true})
    assert_equal(xlsx_styles, {bg_color: '333333', fg_color: '000000', alignment: {horizontal: true}, b: true, sz: 14, i: true, u: true, test: true})
  end

  test "range_hash_to_str" do
    num_columns = 26
    num_rows = 10

    assert_raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError do
      klass.range_hash_to_str({}, num_columns, num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError do
      klass.range_hash_to_str({columns: 1}, num_columns, num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError do
      klass.range_hash_to_str({rows: 1}, num_columns, num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError do
      klass.range_hash_to_str({columns: [1,2,3], rows: 1}, num_columns, num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError do
      klass.range_hash_to_str({columns: 1, rows: [1,2,3]}, num_columns, num_rows)
    end

    assert_equal klass.range_hash_to_str({columns: 0, rows: 1}, num_columns, num_rows), "A1:A1"

    assert_equal klass.range_hash_to_str({columns: (0..2), rows: (1..3)}, num_columns, num_rows), "A1:C3"

    assert_equal klass.range_hash_to_str({columns: ('A'..'C'), rows: (1..3)}, num_columns, num_rows), "A1:C3"

    assert_equal klass.range_hash_to_str({columns: :all, rows: :all}, num_columns, num_rows), "A1:Z10"

    assert_raise SpreadsheetArchitect::Exceptions::BadRangeError do
      assert_equal klass.range_hash_to_str({columns: ('foobar'..'asd'), rows: (1..3)}, num_columns, num_rows), "foobar1:asd3"
    end
  end

  test "verify_range" do
    num_rows = 10

    klass.verify_range("A1:Z10", num_rows)

    assert_raise SpreadsheetArchitect::Exceptions::BadRangeError do
      klass.verify_range(:foo, num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::BadRangeError do
      klass.verify_range("A1Z10", num_rows)
    end

    assert_raise SpreadsheetArchitect::Exceptions::BadRangeError do
      klass.verify_range("A1:A11", num_rows)
    end

    #assert_raise SpreadsheetArchitect::Exceptions::BadRangeError do
      klass.verify_range("A1:AAA1", num_rows)
    #end
  end

  test "verify_column" do
    num_columns = 26

    klass.verify_column(1, num_columns)

    klass.verify_column("A", num_columns)

    assert_raise SpreadsheetArchitect::Exceptions::InvalidColumnError do
      klass.verify_column(-1, num_columns)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidColumnError do
      klass.verify_column(99, num_columns)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidColumnError do
      klass.verify_column("foobar", num_columns)
    end

    assert_raise SpreadsheetArchitect::Exceptions::InvalidColumnError do
      klass.verify_column([], num_columns)
    end

    #assert_raise SpreadsheetArchitect::Exceptions::InvalidColumnError do
      klass.verify_column("ZZ", num_columns)
    #end
  end

  test "symbolize_keys" do
    hash = klass.symbolize_keys
    assert_empty hash

    hash = klass.symbolize_keys({'foo' => :bar})
    assert_nil hash['foo']
    assert_equal hash[:foo], :bar

    hash = klass.symbolize_keys({'foo' => :bar, bar: :foo})
    assert_nil hash['foo']
    assert_equal hash[:foo], :bar

    hash = klass.symbolize_keys({'foo' => {'foo' => :bar}})
    assert_nil hash['foo']
    assert_equal hash[:foo][:foo], :bar

    hash = klass.symbolize_keys({'foo' => {'foo' => {'foo' => :bar}}})
    assert_nil hash['foo']
    assert_equal hash[:foo][:foo][:foo], :bar
  end

  test "constants" do
    assert_equal klass::COL_NAMES.count, 16384
    assert_equal klass::COL_NAMES.first, 'A'
    assert_equal klass::COL_NAMES.last, 'XFD'
  end

end
