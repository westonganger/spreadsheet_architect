# Spreadsheet Architect
<a href='https://ko-fi.com/A5071NK' target='_blank'><img height='32' style='border:0px;height:32px;' src='https://az743702.vo.msecnd.net/cdn/kofi1.png?v=a' border='0' alt='Buy Me a Coffee' /></a> 

Spreadsheet Architect is a library that allows you to create XLSX, ODS, or CSV spreadsheets easily from ActiveRecord relations, Plain Ruby classes, or predefined data.

Key Features:

- Can generate headers & columns from ActiveRecord column_names or a Class/Model's `spreadsheet_columns` method
- Dead simple custom spreadsheets with custom data
- Data Sources: ActiveRecord relations, array of Ruby Objects, or 2D Array Data
- Easily style and customize spreadsheets
- Create multi sheet spreadsheets
- Setting Class/Model or Project specific defaults
- Simple to use ActionController renderers for Rails
- Plain Ruby (without Rails) supported

Spreadsheet Architect adds the following methods:
```ruby
# Rails ActiveRecord Model
Post.order(name: :asc).where(published: true).to_xlsx
Post.order(name: :asc).where(published: true).to_ods
Post.order(name: :asc).where(published: true).to_csv

# Plain Ruby Class
Post.to_xlsx(instances: posts_array)
Post.to_ods(instances: posts_array)
Post.to_csv(instances: posts_array)

# One Time Usage
headers = ['Col 1','Col 2','Col 3']
data = [[1,2,3], [4,5,6], [7,8,9]]
SpreadsheetArchitect.to_xlsx(data: data, headers: headers)
SpreadsheetArchitect.to_ods(data: data, headers: headers)
SpreadsheetArchitect.to_csv(data: data, header: false)
```

# Install
```ruby
# Gemfile
gem 'spreadsheet_architect'
```

# Basic Class/Model Setup

### Model
```ruby
class Post < ActiveRecord::Base #activerecord not required
  include SpreadsheetArchitect

  belongs_to :author
  belongs_to :category
  has_many :tags

  #optional for activerecord classes, defaults to the models column_names
  def spreadsheet_columns

    #[[Label, Method/Statement, Type(optional) to Call on each Instance, Cell Type(optional)]....]
    [
      ['Title', :title],
      ['Content', content],
      ['Author', (author.name if author)],
      ['Published?', (published ? 'Yes' : 'No')],
      ['Published At', :published_at],
      ['# of Views', :number_of_views, :float],
      ['Rating', :rating],
      ['Category/Tags', "#{category.name} - #{tags.collect(&:name).join(', ')}"]
    ]

    # OR if you want to use the method or attribute name as a label it must be a symbol ex. "Title", "Content", "Published"
    [:title, :content, :published]

    # OR a Combination of Both ex. "Title", "Content", "Author Name", "Published"
    [:title, :content, ['Author Name',(author.name rescue nil)], ['# of Views', :number_of_views, :float], :published]
  end
end
```

Note: Do not define your labels inside this method if you are going to be using custom headers in the model or project defaults.

# Usage

### Method 1: Controller (for Rails)
```ruby

class PostsController < ActionController::Base
  respond_to :html, :xlsx, :ods, :csv

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
# Ex. with ActiveRecord relation
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write Post.order(published_at: :asc).to_xlsx
end
File.open('path/to/file.ods', 'w+b') do |f|
  f.write Post.order(published_at: :asc).to_ods
end
File.open('path/to/file.csv', 'w+b') do |f|
  f.write Post.order(published_at: :asc).to_csv
end

# Ex. with plain ruby class
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write Post.to_xlsx(instances: posts_array)
end

# Ex. One time Usage
File.open('path/to/file.xlsx', 'w+b') do |f|
  headers = ['Col 1','Col 2','Col 3']
  data = [[1,2,3], [4,5,6], [7,8,9]]
  f.write SpreadsheetArchitect::to_xlsx(data: data, headers: headers)
end
```
<br>

# Methods & Options


## SomeClass.to_xlsx

|Option|Default|Notes|
|---|---|---|
|**spreadsheet_columns**<br>*Array*| This defaults to your models custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Use this option to override the model instances `spreadsheet_columns` method|
|**instances**<br>*Array*| |**Required only for Non-ActiveRecord classes** Array of class/model instances.|
|**headers**<br>*2D Array*|This defaults to your models custom spreadsheet_columns method or `self.column_names.collect(&:titleize)`|Pass `false` to skip the header row.|
|**sheet_name**<br>*String*|The class name||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false, format_code: nil}`|Styles for non-header rows. See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**column_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**range_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**merges**<br>*Array*||Merge cells. [See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**borders**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**column_widths**<br>*Array*||Sometimes you  may want explicit column widths. Use nil if you want a column to autofit again.|
  
<br>

## SomeClass.to_ods

|Option|Default|Notes|
|---|---|---|
|**spreadsheet_columns**<br>*Array*| This defaults to your models custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Use this option to override the model instances `spreadsheet_columns` method|
|**instances**<br>*Array*| |**Required only for Non-ActiveRecord models** Array of class/model instances.|
|**headers**<br>*2D Array*|`self.column_names.collect(&:titleize)`|Pass `false` to skip the header row.|
|**sheet_name**<br>*String*|The class name||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_size: 10, bold: true}`|Note: Currently ODS only supports these options (values can be changed though)|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently ODS only supports these options|
  
<br>

## SomeClass.to_csv

|Option|Default|Notes|
|---|---|---|
|**spreadsheet_columns**<br>*Array*| This defaults to your models custom `spreadsheet_columns` method or any custom defaults defined.<br>If none of those then falls back to `self.column_names` for ActiveRecord models. | Use this to option override the model instances `spreadsheet_columns` method|
|**instances**<br>*Array*| |**Required only for Non-ActiveRecord classes** Array of class/model instances.|
|**headers**<br>*2D Array*|`self.column_names.collect(&:titleize)`| Data for the header rows cells. Pass `false` to skip the header row.|

<br>

## SpreadsheetArchitect.to_xlsx

|Option|Default|Notes|
|---|---|---|
|**data**<br>*Array*| |Data for the non-header row cells. |
|**headers**<br>*2D Array*|`false`|Data for the header row cells. Pass `false` to skip the header row.|
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false, format_code: nil}`|Styles for non-header rows. See all available style options [here](https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_styles_reference.md)|
|**column_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**range_styles**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**merges**<br>*Array*||Merge cells. [See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**borders**<br>*Array*||[See this example for usage](https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb)|
|**column_types**<br>*Array*||Valid types for XLSX are :string, :integer, :float, :date, :time, :boolean, nil = auto determine|
|**column_widths**<br>*Array*||Sometimes you may want explicit column widths. Use nil if you want a column to autofit again.|
  
<br> 

## SpreadsheetArchitect.to_ods

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Data for the non-header row cells.|
|**headers**<br>*2D Array*|`false`|Data for the header rows cells. Pass `false` to skip the header row.|
|**sheet_name**<br>*String*|`Sheet1`||
|**header_style**<br>*Hash*|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_size: 10, bold: true}`|Note: Currently ODS only supports these options|
|**row_style**<br>*Hash*|`{background_color: nil, color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently ODS only supports these options|
|**column_types**<br>*Array*||Valid types for ODS are :string, :float, :date, :percent, :currency, nil = auto determine|
  
<br>

## SpreadsheetArchitect.to_csv

|Option|Default|Notes|
|---|---|---|
|**data**<br>*2D Array*| |Data for the non-header row cells.|
|**headers**<br>*2D Array*|`false`|Data for the header rows cells. Pass `false` to skip the header row.|


# Change model default method options
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
    column_types: []
  }
end
```

# Change project wide default method options
```ruby
# config/initializers/spreadsheet_architect.rb

SpreadsheetArchitect.default_options = {
  headers: true,
  header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
  sheet_name: 'My Project Export',
  column_styles: [],
  range_styles: [],
  merges: [],
  borders: [],
  column_types: []
}
```

# Complex XLSX Example with Styling
See this example: https://github.com/westonganger/spreadsheet_architect/blob/master/examples/complex_xlsx_styling.rb


# Multi Sheet XLSX or ODS spreadsheets
```ruby
# Returns corresponding spreadsheet libraries object
package = SpreadsheetArchitect.to_axlsx_package({data: data, headers: headers})
SpreadsheetArchitect.to_axlsx_package({data: data, headers: headers}, package) # to combine two sheets to one file

spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({data: data, headers: headers})
SpreadsheetArchitect.to_rodf_spreadsheet({data: data, headers: headers}, spreadsheet) # to combine two sheets to one file
```

See this example: https://github.com/westonganger/spreadsheet_architect/blob/master/examples/multi_sheet_spreadsheets.rb


# Axlsx Style Reference
I have compiled a list of all available style options for axlsx here: https://github.com/westonganger/spreadsheet_architect/blob/master/docs/axlsx_style_reference.md


# Credits
Created by [@westonganger](https://github.com/westonganger)

For any consulting or contract work please contact me via my company website: [Solid Foundation Web Development](https://solidfoundationwebdev.com)

[![Solid Foundation Web Development Logo](https://solidfoundationwebdev.com/logo-sm.png)](https://solidfoundationwebdev.com)
