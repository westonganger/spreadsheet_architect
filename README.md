# Spreadsheet Architect

<a href="https://badge.fury.io/rb/spreadsheet_architect" target="_blank"><img height="21" style='border:0px;height:21px;' border='0' src="https://badge.fury.io/rb/spreadsheet_architect.svg" alt="Gem Version"></a>
<a href='https://travis-ci.com/westonganger/spreadsheet_architect' target='_blank'><img height='21' style='border:0px;height:21px;' src='https://api.travis-ci.org/westonganger/spreadsheet_architect.svg?branch=master' border='0' alt='Build Status' /></a>
<a href='https://rubygems.org/gems/spreadsheet_architect' target='_blank'><img height='21' style='border:0px;height:21px;' src='https://ruby-gem-downloads-badge.herokuapp.com/spreadsheet_architect?label=rubygems&type=total&total_label=downloads&color=brightgreen' border='0' alt='RubyGems Downloads' /></a>
<a href='https://ko-fi.com/A5071NK' target='_blank'><img height='22' style='border:0px;height:22px;' src='https://az743702.vo.msecnd.net/cdn/kofi1.png?v=a' border='0' alt='Buy Me a Coffee' /></a>

Spreadsheet Architect is a library that allows you to create XLSX, ODS, or CSV spreadsheets super easily from ActiveRecord relations, plain Ruby objects, or tabular data.

Key Features:

- Dead simple custom spreadsheets with custom data
- Data Sources: Tabular Data from an Array, ActiveRecord relations, or array of plain Ruby object instances
- Easily style and customize spreadsheets
- Create multi sheet spreadsheets
- Setting Class/Model or Project specific defaults
- Simple to use ActionController renderers for Rails
- Plain Ruby (without Rails) completely supported

# Install
```ruby
gem 'spreadsheet_architect'
```

# General Usage

### Tabular (Array) Data

The simplest and preffered usage is to simply create the data array yourself.

```ruby
headers = ['Col 1','Col 2','Col 3']
data = [[1,2,3], [4,5,6], [7,8,9]]
SpreadsheetArchitect.to_xlsx(headers: headers, data: data)
SpreadsheetArchitect.to_ods(headers: headers, data: data)
SpreadsheetArchitect.to_csv(headers: headers, data: data)
```

Using this style will allow you to utilize any custom performance optimizations during your data generation process. This will come in handy when the spreadsheets get large and things start to get slow. One of my favourites for Rails is [light_record](https://github.com/Paxa/light_record)

### Rails Relations or an Array of plain Ruby object instances

If you would like to add the methods `to_xlsx`, `to_ods`, `to_csv`, `to_axlsx_package`, `to_rodf_spreadsheet` to some class, you can simply include the SpreadsheetArchitect module to whichever classes you choose. For example:

```ruby
class Post < ApplicationRecord
  include SpreadsheetArchitect
end
```

When using on an AR Relation or using the `:instances` option, SpreadsheetArchitect requires an instance method to be defined on the class to generate the data. By default it looks for the `spreadsheet_columns` method on the class. If you are using on an ActiveRecord model and that method is not defined, it would fallback to the models `column_names` method. If using the `:data` option this is completely ignored.

```ruby
class Post
  include SpreadsheetArchitect

  def spreadsheet_columns
    ### Column format is: [Header, Cell Data / Method (if symbol) to Call on each Instance, (optional) Cell Type]
    [
      ['Title', :title],
      ['Content', content.strip],
      ['Author', (author.name if author)],
      ['Published?', (published ? 'Yes' : 'No')],
      :published_at, # uses the method name as header title Ex. 'Published At'
      ['# of Views', :number_of_views, :float],
      ['Rating', :rating],
      ['Category/Tags', "#{category.name} - #{tags.collect(&:name).join(', ')}"]
    ]
  end

end
```

Then use it on the class or ActiveRecord relations of the class

```ruby
posts = Post.order(name: :asc).where(published: true)
posts.to_xlsx
posts.to_ods
posts.to_csv

# Plain Ruby Objects
posts_array = 10.times.map{|i| Post.new(number: i)}
Post.to_xlsx(instances: posts_array)
Post.to_ods(instances: posts_array)
Post.to_csv(instances: posts_array)
```

If you want to use a different method name then `spreadsheet_columns` you can pass a method name to the `:spreadsheet_columns` option.

```ruby
Post.to_xlsx(instances: posts, spreadsheet_columns: :my_special_method)
```

Alternatively, you can pass a Proc/lambda to the `spreadsheet_columns` option. For those purists that really dont want to define any extra `spreadsheet_columns` instance method on your model, this option can help you work with that methodology.

```ruby
Post.to_xlsx(instances: posts, spreadsheet_columns: Proc.new{|instance|
  [
    ['Title', :title],
    ['Content', instance.content.strip],
    ['Author', (instance.author.name if instance.author)],
    ['Published?', (instance.published ? 'Yes' : 'No')],
    :published_at, # uses the method name as header title Einstance. 'Published At'
    ['# of Views', :number_of_views, :float],
    ['Rating', :rating],
    ['Category/Tags', "#{instance.category.name} - #{instance.tags.collect(&:name).join(', ')}"]
  ]
})
```

# Sending & Saving Spreadsheets

### Method 1: Save to a file manually

```ruby
file_data = SpreadsheetArchitect.to_xlsx(headers: headers, data: data)

File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write file_data
end
```

### Method 2: Send Data via Rails Controller

```ruby
class PostsController < ActionController::Base
  respond_to :html, :xlsx, :ods, :csv

  def index
    @posts = Post.order(published_at: :asc)

    render xlsx: @posts
  end

  # Using respond_with
  def index
    @posts = Post.order(published_at: :asc)

    respond_with @posts
  end

  # OR Using respond_with with custom options
  def index
    @posts = Post.order(published_at: :asc)

    if ['xlsx','ods','csv'].include?(request.format)
      respond_with @posts.to_xlsx(row_style: {bold: true}), filename: 'Posts'
    else
      respond_with @posts
    end
  end

  # OR Using responders
  def index
    @posts = Post.order(published_at: :asc)

    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts }
      format.ods { render ods: @posts }
      format.csv{ render csv: @posts }
    end
  end

  # OR Using responders with custom options
  def index
    @posts = Post.order(published_at: :asc)

    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts.to_xlsx(headers: false) }
      format.ods { render ods: Post.to_ods(instances: @posts) }
      format.csv{ render csv: @posts.to_csv(headers: false), file_name: 'articles' }
    end
  end
end
```

# Multi Sheet Spreadsheets

### XLSX

```ruby
axlsx_package = SpreadsheetArchitect.to_axlsx_package({headers: headers, data: data})
axlsx_package = SpreadsheetArchitect.to_axlsx_package({headers: headers, data: data}, axlsx_package)

File.open('path/to/multi_sheet_file.xlsx', 'w+b') do |f|
  f.write axlsx_package.to_stream.read
end
```

See this file for more details: https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/multi_sheet_test.rb

### ODS
```ruby
ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({headers: headers, data: data})
ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({headers: headers, data: data}, ods_spreadsheet)

File.open('path/to/multi_sheet_file.ods', 'w+b') do |f|
  f.write ods_spreadsheet
end
```

See this file for more details: https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/multi_sheet_test.rb

# Methods

## `to_xlsx(options={})`

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Cannot be used with the `:instances` option.<br><br>Tabular data for the non-header row cells.  |
|**instances**<br>*Array*| |Cannot be used with the `:data` option.<br><br>Array of class/model instances to be used as row data. Cannot be used with :data option|
|**spreadsheet_columns**<br>*Proc/Symbol/String*| Use this option to override or define the spreadsheet columns. Normally, if this option is not specified and are using the instances option/ActiveRecord relation, it uses the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If neither of those and is an ActiveRecord model, then it will falls back to the models `self.column_names` | Cannot be used with the `:data` option.<br><br>If a Proc value is passed it will be evaluated on the instance object.<br><br>If a Symbol or String value is passed then it will search the instance for a method name that matches and call it. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_style_reference.md)|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false, format_code: nil}`|Styles for non-header rows. See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_style_reference.md)|
|**column_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb)|
|**range_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb)|
|**conditional_row_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb). The if/unless proc will called with the following args: `row_index`, `row_data`|
|**merges**<br>*Array*||Merge cells. [See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb). Warning merges cannot overlap eachother, if you attempt to do so Excel will claim your spreadsheet is corrupt and refuse to open your spreadsheet.|
|**borders**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb)|
|**column_types**<br>*Array*||Valid types for XLSX are :string, :integer, :float, :date, :time, :boolean, nil = auto determine.|
|**column_widths**<br>*Array*||Sometimes you may want explicit column widths. Use nil if you want a column to autofit again.|
|**freeze_headers**<br>*Boolean*||Make all header rows frozen/fixed so they do not scroll.|
|**freeze**<br>* Hash*|`{rows: (1..4), columns: :all}`|Make all specified rows and columns frozen/fixed so they do not scroll.|

## `to_axlsx_spreadsheet(options={}, axlsx_package_to_join=nil)`
Same options as `to_xlsx`

## `to_ods(options={})`

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Cannot be used with the `:instances` option.<br><br>Tabular data for the non-header row cells.  |
|**instances**<br>*Array*| |Cannot be used with the `:data` option.<br><br>Array of class/model instances to be used as row data. Cannot be used with :data option|
|**spreadsheet_columns**<br>*Proc/Symbol/String*| Use this option to override or define the spreadsheet columns. Normally, if this option is not specified and are using the instances option/ActiveRecord relation, it uses the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If neither of those and is an ActiveRecord model, then it will falls back to the models `self.column_names` | Cannot be used with the `:data` option.<br><br>If a Proc value is passed it will be evaluated on the instance object.<br><br>If a Symbol or String value is passed then it will search the instance for a method name that matches and call it. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_size: 10, bold: true}`|Note: Currently ODS only supports these options|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently ODS only supports these options|
|**column_types**<br>*Array*||Valid types for ODS are :string, :float, :date, :time, :boolean, nil = auto determine. Due to [RODF Issue #19](https://github.com/thiagoarrais/rodf/issues/19), :date/:time will be converted to :string |

## `to_rodf_spreadsheet(options={}, spreadsheet_to_join=nil)`
Same options as `to_ods`

## `to_csv(options={})`

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Cannot be used with the `:instances` option.<br><br>Tabular data for the non-header row cells.  |
|**instances**<br>*Array*| |Cannot be used with the `:data` option.<br><br>Array of class/model instances to be used as row data. Cannot be used with :data option|
|**spreadsheet_columns**<br>*Proc/Symbol/String*| Use this option to override or define the spreadsheet columns. Normally, if this option is not specified and are using the instances option/ActiveRecord relation, it uses the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If neither of those and is an ActiveRecord model, then it will falls back to the models `self.column_names` | Cannot be used with the `:data` option.<br><br>If a Proc value is passed it will be evaluated on the instance object.<br><br>If a Symbol or String value is passed then it will search the instance for a method name that matches and call it. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |


# Change class-wide default method options

```ruby
class Post < ApplicationRecord
  include SpreadsheetArchitect

  def spreadsheet_columns
    [:name, :content]
  end

  SPREADSHEET_OPTIONS = {
    headers: [
      ['My Post Report'],
      self.column_names.map{|x| x.titleize}
    ],
     spreadsheet_columns: :spreadsheet_columns,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: self.name,
    column_styles: [],
    range_styles: [],
    conditional_row_styles: [],
    merges: [],
    borders: [],
    column_types: [],
  }
end
```

# Change project-wide default method options

```ruby
# config/initializers/spreadsheet_architect.rb

SpreadsheetArchitect.default_options = {
  headers: true,
  spreadsheet_columns: :spreadsheet_columns,
  header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  sheet_name: 'My Project Export',
  column_styles: [],
  range_styles: [],
  conditional_row_styles: [],
  merges: [],
  borders: [],
  column_types: [],
}
```

# Kitchen Sink Examples with Styling for XLSX and ODS
See this example: https://github.com/westonganger/spreadsheet_architect/blob/master/test/unit/kitchen_sink_test.rb

# Axlsx Style Reference

I have compiled a list of all available style options for axlsx here: https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_style_reference.md

# Testing / Validating your Spreadsheets

A wise word of advice, when testing your spreadsheets I recommend to use Excel instead of LibreOffice. This is because I have seen through testing, that where LibreOffice seems to just let most incorrect things just slide on through, Excel will not even open the spreadsheet as apparently it is much more strict about the spreadsheet validations. This will help you better identify any incorrect styling or customization issues.

# Contributing

We use the `appraisal` gem for testing multiple versions of `axlsx`. Please use the following steps to test using `appraisal`.

1. `bundle exec appraisal install`
2. `bundle exec appraisal rake test`

At this time the spreadsheets generated by the test suite are manually inspected. After running the tests, the test output can be viewed at `tmp/#{alxsx_version}/*`

# Credits

Created & Maintained by [Weston Ganger](https://westonganger.com) - [@westonganger](https://github.com/westonganger)

For any consulting or contract work please contact me via my company website: [Solid Foundation Web Development](https://solidfoundationwebdev.com)

[![Solid Foundation Web Development Logo](https://solidfoundationwebdev.com/logo-sm.png)](https://solidfoundationwebdev.com)
