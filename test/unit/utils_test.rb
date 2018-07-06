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
    assert klass.get_cell_data(@options, SpreadsheetArchitect)

    assert klass.get_cell_data(@options, Post)

    assert klass.get_options({}, SpreadsheetArchitect)

    assert_raise SpreadsheetArchitect::Exceptions::MultipleDataSourcesError do
      klass.get_cell_data(@options.merge(instances: []), SpreadsheetArchitect)
    end

    ### using Data option
    output = klass.get_cell_data(@options.merge(headers: true), SpreadsheetArchitect)
    assert_equal [[]], output[:headers]

    output = klass.get_cell_data(@options.merge(column_types: nil), SpreadsheetArchitect)
    assert_nil output[:column_types]

    output = klass.get_cell_data(@options.merge(column_types: []), SpreadsheetArchitect)
    assert_nil output[:column_types]

    output = klass.get_cell_data(@options.merge(column_types: [:string]), SpreadsheetArchitect)
    assert_equal output[:column_types], [:string]

    headers = [[1,2,3], [3,4,5]]
    output = klass.get_cell_data(@options.merge(headers: headers), SpreadsheetArchitect)
    assert_equal headers, output[:headers]

    ### Using instances option
    output = klass.get_cell_data(@options.merge(data: nil), Post.all)
    assert output[:instances].is_a?(Array)

    output = klass.get_cell_data(@options.merge(data: nil), Post.limit(0))
    assert output[:instances].is_a?(Array)

    assert_raise SpreadsheetArchitect::Exceptions::NoDataError do
      klass.get_cell_data(@options.merge(data: nil, instances: nil), SpreadsheetArchitect)
    end

    output = klass.get_cell_data(@options.merge(data: nil, instances: [PlainRubyObject.new]), SpreadsheetArchitect)
    assert output[:instances].count == 1
  end

  test "get_options" do
    ### Empty
    assert_not_empty klass.get_options({}, SpreadsheetArchitect)

    ### using SpreadsheetArchitect
    assert_not_empty klass.get_options(@options, SpreadsheetArchitect)

    ### with model defaults via SPREADSHEET_OPTIONS 
    assert defined?(CustomPost::SPREADSHEET_OPTIONS)
    assert klass.get_options(@options, CustomPost)

    ### without model defaults via SPREADSHEET_OPTIONS 
    assert_not_empty klass.get_options(@options, Post)

    ### without :headers removes :header_style
    assert_equal klass.get_options({header_style: false, headers: false}, SpreadsheetArchitect)[:header_style], false

    ### sets :sheet_name if needed 
    assert_equal klass.get_options({sheet_name: false}, SpreadsheetArchitect)[:sheet_name], 'Sheet1'

    ### sets :sheet_name if needed, using pluralized only when using Rails
    assert_equal klass.get_options({sheet_name: false}, Post)[:sheet_name], 'Posts'
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

  test "is_ar_model" do
    assert klass.is_ar_model?(Post)

    assert_not klass.is_ar_model?(SpreadsheetArchitect)
  end

  test "str_titleize" do
    assert_equal(klass.str_titleize('my_project_export'), 'My Project Export')
    assert_equal(klass.str_titleize('My Project Export'), 'My Project Export')
    assert_equal(klass.str_titleize('TBS report'), 'TBS Report')
  end

  test "check_option_type" do
    klass.check_option_type(@options, :data, Array)

    klass.check_option_type(@options, :foo, Array)

    assert_raise SpreadsheetArchitect::Exceptions::InvalidTypeError do
      klass.check_option_type({foo: :bar}, :foo, Array)
    end
  end

  test "verify_option_types" do
    klass.verify_option_types(@options)

    assert_raise SpreadsheetArchitect::Exceptions::InvalidTypeError do
      klass.verify_option_types(@options.merge({column_widths: :foobar}))
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
