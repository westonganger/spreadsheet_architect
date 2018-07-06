# Spreadsheet Architect

<a href="https://badge.fury.io/rb/spreadsheet_architect" target="_blank"><img height="21" style='border:0px;height:21px;' border='0' src="https://badge.fury.io/rb/spreadsheet_architect.svg" alt="Gem Version"></a>
<a href='https://travis-ci.org/westonganger/spreadsheet_architect' target='_blank'><img height='21' style='border:0px;height:21px;' src='https://api.travis-ci.org/westonganger/spreadsheet_architect.svg?branch=master' border='0' alt='Build Status' /></a>
<a href='https://rubygems.org/gems/spreadsheet_architect' target='_blank'><img height='21' style='border:0px;height:21px;' src='https://ruby-gem-downloads-badge.herokuapp.com/spreadsheet_architect?label=rubygems&type=total&total_label=downloads&color=brightgreen' border='0' alt='RubyGems Downloads' /></a>
<a href='https://ko-fi.com/A5071NK' target='_blank'><img height='22' style='border:0px;height:22px;' src='https://az743702.vo.msecnd.net/cdn/kofi1.png?v=a' border='0' alt='Buy Me a Coffee' /></a> 

Spreadsheet Architect is a library that allows you to create XLSX, ODS, or CSV spreadsheets super easily from ActiveRecord relations, plain Ruby objects, or tabular data.

Key Features:

- Dead simple custom spreadsheets with custom data
- Data Sources: ActiveRecord relations, array of plain Ruby object instances, or tabular 2D Array Data
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

```ruby
headers = ['Col 1','Col 2','Col 3']
data = [[1,2,3], [4,5,6], [7,8,9]]
SpreadsheetArchitect.to_xlsx(headers: headers, data: data)
SpreadsheetArchitect.to_ods(headers: headers, data: data)
SpreadsheetArchitect.to_csv(headers: headers, data: data)
```

### Rails relation or an array of plain Ruby object instances

```ruby
posts = Post.order(name: :asc).where(published: true)
# OR
posts = 10.times.map{|i| Post.new(number: i)}

SpreadsheetArchitect.to_xlsx(instances: posts)
SpreadsheetArchitect.to_ods(instances: posts)
SpreadsheetArchitect.to_csv(instances: posts)
```

**(Optional)** If you would like to add the methods `to_xlsx`, `to_ods`, `to_csv`, `to_axlsx_package`, `to_rodf_spreadsheet` to some class, you can simply include the SpreadsheetArchitect module to whichever classes you choose. A good default strategy is to simply add it to the ApplicationRecord or another parent class to have it available on all appropriate classes. For example:

```ruby
class ApplicationRecord < ActiveRecord::Base
  include SpreadsheetArchitect
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

# Usage with Instances / ActiveRecord Relations

When NOT using the `:data` option, ie. on an AR Relation or using the `:instances` option, Spreadsheet Architect requires an instance method defined on the class to generate the data. It looks for the `spreadsheet_columns` method on the class. If you are using on an ActiveRecord model and that method is not defined, it would fallback to the models `column_names` method (not recommended). If using the `:data` option this is ignored.

```ruby
class Post

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
```

Alternatively, if `spreadsheet_columns` is passed as an option, this instance method does not need to be defined on the class. If defined on the class then naturally this will override it.

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

### Method 1: Send Data via Rails Controller

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

### Method 2: Save to a file manually

```ruby
### Ex. with ActiveRecord relation
file_data = Post.order(published_at: :asc).to_xlsx
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write file_data
end

file_data = Post.order(published_at: :asc).to_ods
File.open('path/to/file.ods', 'w+b') do |f|
  f.write file_data
end

file_data = Post.order(published_at: :asc).to_csv
File.open('path/to/file.csv', 'w+b') do |f|
  f.write file_data
end
```

# Multi Sheet XLSX Spreadsheets
```ruby
axlsx_package = SpreadsheetArchitect.to_axlsx_package({headers: headers, data: data})
axlsx_package = SpreadsheetArchitect.to_axlsx_package({headers: headers, data: data}, package)

File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write axlsx_package.to_stream.read
end
```

See this file for more details: https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/multi_sheet_test.rb

### Multi Sheet ODS Spreadsheets
```ruby
ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({headers: headers, data: data})
ods_spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({headers: headers, data: data}, spreadsheet)

File.open('path/to/file.ods', 'w+b') do |f|
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
|**spreadsheet_columns**<br>*Array*| If using the instances option or on a ActiveRecord relation, this defaults to the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Cannot be used with the `:data` option.<br><br>Use this option to override or define the spreadsheet columns. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false, format_code: nil}`|Styles for non-header rows. See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**column_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/kitchen_sink_test.rb)|
|**range_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/kitchen_sink_test.rb)|
|**merges**<br>*Array*||Merge cells. [See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/kitchen_sink_test.rb)|
|**borders**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/kitchen_sink_test.rb)|
|**column_types**<br>*Array*||Valid types for XLSX are :string, :integer, :float, :boolean, nil = auto determine.|
|**column_widths**<br>*Array*||Sometimes you may want explicit column widths. Use nil if you want a column to autofit again.|

## `to_axlsx_spreadsheet(options={}, axlsx_package_to_join=nil)`
Same options as `to_xlsx`. For more details

## `to_ods(options={})`

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Cannot be used with the `:instances` option.<br><br>Tabular data for the non-header row cells.  |
|**instances**<br>*Array*| |Cannot be used with the `:data` option.<br><br>Array of class/model instances to be used as row data. Cannot be used with :data option|
|**spreadsheet_columns**<br>*Array*| If using the instances option or on a ActiveRecord relation, this defaults to the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Cannot be used with the `:data` option.<br><br>Use this option to override or define the spreadsheet columns. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_size: 10, bold: true}`|Note: Currently ODS only supports these options|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently ODS only supports these options|
|**column_types**<br>*Array*||Valid types for ODS are :string, :float :percent, :currency, :date, :time,, nil = auto determine. Due to [RODF Issue #19](https://github.com/thiagoarrais/rodf/issues/19), :date/:time will be converted to :string |

## `to_rodf_spreadsheet(options={}, spreadsheet_to_join=nil)`
Same options as `to_ods`

## `to_csv(options={})`

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Cannot be used with the `:instances` option.<br><br>Tabular data for the non-header row cells.  |
|**instances**<br>*Array*| |Cannot be used with the `:data` option.<br><br>Array of class/model instances to be used as row data. Cannot be used with :data option|
|**spreadsheet_columns**<br>*Array*| If using the instances option or on a ActiveRecord relation, this defaults to the classes custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Cannot be used with the `:data` option.<br><br>Use this option to override or define the spreadsheet columns. |
|**headers**<br>*Array / 2D Array*| |Data for the header row cells. If using on a class/relation, this defaults to the ones provided via `spreadsheet_columns`. Pass `false` to skip the header row. |


# Change class-wide default method options

```ruby
class Post
  include SpreadsheetArchitect

  def spreadsheet_columns
    [:name, :content]
  end

  SPREADSHEET_OPTIONS = {
    headers: [
      ['My Post Report'],
      self.column_names.map{|x| x.titleize}
    ],
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: self.name,
    column_styles: [],
    range_styles: [],
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
  header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  sheet_name: 'My Project Export',
  column_styles: [],
  range_styles: [],
  merges: [],
  borders: [],
  column_types: [],
}
```

# Kitchen Sink Examples with Styling for XLSX and ODS
See this example: https://github.com/westonganger/spreadsheet_architect/blob/master/test/spreadsheet_architect/kitchen_sink_test.rb

# Axlsx Style Reference

I have compiled a list of all available style options for axlsx here: https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_style_reference.md

# Contributing

We use the `appraisal` gem for testing multiple versions of `axlsx`. Please use the following steps to test using `appraisal`.

1. `bundle exec appraisal install`
2. `bundle exec appraisal rake test`

# Credits

Created & Maintained by [Weston Ganger](https://westonganger.com) - [@westonganger](https://github.com/westonganger)

For any consulting or contract work please contact me via my company website: [Solid Foundation Web Development](https://solidfoundationwebdev.com)

[![Solid Foundation Web Development Logo](https://solidfoundationwebdev.com/logo-sm.png)](https://solidfoundationwebdev.com)
