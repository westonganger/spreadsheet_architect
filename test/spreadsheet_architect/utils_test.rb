require 'test_helper'

class UtilsTest < ActiveSupport::TestCase
  klass = SpreadsheetArchitect::Utils

  def setup
    @options = {
      header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      sheet_name: 'My Project Export',
      column_styles: [],
      range_styles: [],
      merges: [],
      borders: [],
      column_types: [],
      column_widths: [],

      headers: ['test1', 'test2','test3'],
      data: [
        ['row1'],
        ['row2 c1', 'row2 c2'],
        ['the','data']
      ]
    }
  end

  def teardown
  end

  test "get_cell_data" do
    assert_not_empty klass.get_cell_data(@options, SpreadsheetArchitect)

    assert_not_empty klass.get_cell_data(@options, Post)

    assert_not_empty klass.get_options({}, SpreadsheetArchitect)

    skip("TODO") # step through all permutations of this method
  end

  test "get_options" do
    assert_not_empty klass.get_options(@options, SpreadsheetArchitect)

    assert_not_empty klass.get_options(@options, Post)

    assert_not_empty klass.get_options({}, SpreadsheetArchitect)

    skip("TODO") # step through all permutations of this method
  end
  
  test "convert_styles_to_ods" do
    ods_styles = klass.convert_styles_to_ods({
      background_color: '333333', 
      color: '000000', 
      align: true, 
      bold: true, 
      font_size: 14, 
      italic: true, 
      underline: true, 
      test: true
    })

    assert_equal(ods_styles, {
      'cell' => {
        'background-color' => '#333333'
      }, 
      'text' => {
        'color' => '#000000', 
        'align' => true, 
        'font-weight' => 'bold', 
        'font-size' => 14, 
        'font-style' => 'italic', 
        'text-underline-type' => 'single', 
        'text-underline-style' => 'solid'
      }
    })
  end

  test "deep_clone" do
    assert_nil klass.deep_clone(nil)
    assert klass.deep_clone('')
    assert klass.deep_clone(1)
    assert klass.deep_clone(1.0)
    assert klass.deep_clone([])
    assert klass.deep_clone({})
    assert klass.deep_clone({foo: :bar})
    assert klass.deep_clone({foo: {foo: :bar}})
  end

  test "is_ar_model" do
    assert klass.is_ar_model?(Post)

    assert_not klass.is_ar_model?(SpreadsheetArchitect)
  end

  test "str_humanize" do
    assert_equal(klass.str_humanize('my_project_export'), 'My Project Export')
    assert_equal(klass.str_humanize('My Project Export'), 'My Project Export')
    assert_equal(klass.str_humanize('TBS report'), 'TBS Report')
  end

  test "check_type" do
    klass.check_type(@options, :data, Array)

    klass.check_type(@options, :foo, Array)

    assert_raise SpreadsheetArchitect::Exceptions::IncorrectTypeError do
      klass.check_type({foo: :bar}, :foo, Array)
    end
  end

  test "check_options_types" do
    klass.check_options_types(@options)

    assert_raise SpreadsheetArchitect::Exceptions::IncorrectTypeError do
      klass.check_options_types(@options.merge({column_widths: :foobar}))
    end
  end

  test "stringify_keys" do
    hash = klass.stringify_keys
    assert_empty hash

    hash = klass.stringify_keys({foo: :bar})
    assert_nil hash[:foo]
    assert_equal hash['foo'], :bar

    hash = klass.stringify_keys({foo: :bar, 'bar' => :foo})
    assert_nil hash[:foo]
    assert_equal hash['foo'], :bar

    hash = klass.stringify_keys({foo: {foo: :bar}})
    assert_nil hash[:foo]
    assert_equal hash['foo']['foo'], :bar

    hash = klass.stringify_keys({foo: {foo: {foo: :bar}}})
    assert_nil hash[:foo]
    assert_equal hash['foo']['foo']['foo'], :bar

    hash = klass.stringify_keys({foo: {foo: {foo: {foo: :bar}}}})
    assert_nil hash[:foo]
    assert_equal hash['foo']['foo']['foo']['foo'], :bar
  end
end
