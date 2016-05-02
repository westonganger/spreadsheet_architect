# Spreadsheet Architect
<a href="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=VKY8YAWAS5XRQ&lc=CA&item_name=Weston%20Ganger&item_number=spreadsheet_architect&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHostedGuest" target="_blank" title="Donate"><img src="https://www.paypalobjects.com/en_US/i/btn/btn_donate_SM.gif" alt="Donate"/></a>

Spreadsheet Architect is a library that allows you to create XLSX, ODS, or CSV spreadsheets easily from ActiveRecord relations, Plain Ruby classes, or predefined data.

Key Features:

- Can generate headers & columns from ActiveRecord column_names, or a Class/Model's spreadsheet_columns method, or one creation with 2D array of data
- Plain Ruby support
- Plain from ActiveRecord relations or Ruby Objects from models ActiveRecord, or 2d Array Data
- Easily style headers and rows
- Model/Class or Project specific defaults
- Simple to use ActionController renderers

Spreadsheet Architect adds the following methods to your class:
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
gem install spreadsheet_architect
```


# Class/Model Setup

### Model
```ruby
class Post < ActiveRecord::Base #activerecord not required
  include SpreadsheetArchitect

  belongs_to :author
  belongs_to :category
  has_many :tags

  #optional for activerecord classes, defaults to the models column_names
  def spreadsheet_columns

    #[[Label, Method/Statement to Call on each Instance, Cell Type(optional)]....]
    [
      ['Title', :title],
      ['Content', content],
      ['Author', (author.name rescue nil)],
      ['Published?', (published ? 'Yes' : 'No')],
      ['Published At', :published_at],
      ['# of Views', :number_of_views],
      ['Rating', :rating],
      ['Category/Tags', "#{category.name} - #{tags.collect(&:name).join(', ')}"]
    ]

    # OR if you want to use the method or attribute name as a label it must be a symbol ex. "Title", "Content", "Published"
    [:title, :content, :published]

    # OR a Combination of Both ex. "Title", "Content", "Author Name", "Published"
    [:title, :content, ['Author Name',(author.name rescue nil)], :published]
  end
end
```

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

  # Using respond_with with custom options
  def index
    @posts = Post.order(published_at: :asc)

    if ['xlsx','ods','csv'].include?(request.format)
      respond_with @posts.to_xlsx(row_style: {bold: true}), filename: 'Posts'
    else
      respond_with @posts
    end
  end

  # Using responders
  def index
    @posts = Post.order(published_at: :asc)

    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts }
      format.ods { render ods: @posts }
      format.csv{ render csv: @posts }
    end
  end

  # Using responders with custom options
  def index
    @posts = Post.order(published_at: :asc)

    respond_to do |format|
      format.html
      format.xlsx { render xlsx: @posts.to_xlsx(headers: false) }
      format.ods { render ods: Post.to_odf(instances: @posts) }
      format.csv{ render csv: @posts.to_csv(headers: false), file_name: 'articles' }
    end
  end
end
```

### Method 2: Save to a file manually
```ruby
# Ex. with ActiveRecord realtion
File.open('path/to/file.xlsx') do |f|
  f.write{ Post.order(published_at: :asc).to_xlsx }
end
File.open('path/to/file.ods') do |f|
  f.write{ Post.order(published_at: :asc).to_ods }
end
File.open('path/to/file.csv') do |f|
  f.write{ Post.order(published_at: :asc).to_csv }
end

# Ex. with plain ruby class
File.open('path/to/file.xlsx') do |f|
  f.write{ Post.to_xlsx(instances: posts_array) }
end

# Ex. One time Usage
File.open('path/to/file.xlsx') do |f|
  headers = ['Col 1','Col 2','Col 3']
  data = [[1,2,3], [4,5,6], [7,8,9]]
  f.write{ SpreadsheetArchitect::to_xlsx(data: data, headers: headers) }
end
```

# Method & Options

<br>
#### `.to_xlsx` - (on custom class/model)
|Option|Type|Default|Notes|
|---|---|---|---|
|**spreadsheet_columns**|Array| AR Model column_names | Required if `spreadsheet_columns` not defined on class except with ActiveRecord models which default to the `column_names` method. Will override models `spreadsheet_columns` method |
|**instances**|Array| |**Required for Non-ActiveRecord classes** Array of class/model instances.|
|**headers**|Boolean|`true`|Pass false to skip the header row.|
|**sheet_name**|String|Class name||
|**header_style**|Hash|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`||
|**row_style**|Hash|`{background_color: nil, color: "FFFFFF", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|Styles for non-header rows.|
  
<br>
#### `.to_ods` - (on custom class/model)
|Option|Type|Default|Notes|
|---|---|---|---|
|**spreadsheet_columns**|Array| AR Model column_names | Required if `spreadsheet_columns` not defined on class except with ActiveRecord models which default to the `column_names` method. Will override models `spreadsheet_columns` method |
|**instances**|Array| |**Required for Non-ActiveRecord classes** Array of class/model instances.|
|**headers**|Boolean|`true`|Pass false to skip the header row.|
|**sheet_name**|String|Class name||
|**header_style**|Hash|`{color: "000000", align: :center, font_size: 10, bold: true}`|Note: Currently only supports these options (values can be changed though)|
|**row_style**|Hash|`{color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently only supports these options (values can be changed though)|
  
<br>
#### `.to_csv` - (on custom class/model)
|Option|Type|Default|Notes|
|---|---|---|---|
|**spreadsheet_columns**|Array| AR Model column_names | Required if `spreadsheet_columns` not defined on class except with ActiveRecord models which default to the `column_names` method. Will override models `spreadsheet_columns` method |
|**instances**|Array| |**Required for Non-ActiveRecord classes** Array of class/model instances.|
|**headers**|Boolean|`true`|Pass false to skip the header row.|

<br>
#### `SpreadsheetArchitect.to_xlsx`
|Option|Type|Default|Notes|
|---|---|---|---|
|**data**|Array| |**Required** 2D Array of data for the non-header row cells. |
|**headers**|Array|`false`|2D Array of data for the header rows cells. Pass false to skip the header row.|
|**sheet_name**|String|`SpreadsheetArchitect`||
|**header_style**|Hash|`{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`||
|**row_style**|Hash|`{background_color: nil, color: "FFFFFF", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`|Styles for non-header rows.|
  
<br> 
#### `SpreadsheetArchitect.to_ods`
|Option|Type|Default|Notes|
|---|---|---|---|
|**data**|Array| |**Required** 2D Array of data for the non-header row cells.|
|**headers**|Array|`false`|2D Array of data for the header rows cells. Pass false to skip the header row.|
|**sheet_name**|String|`SpreadsheetArchitect`||
|**header_style**|Hash|`{color: "000000", align: :center, font_size: 10, bold: true}`|Note: Currently only supports these options (values can be changed though)|
|**row_style**|Hash|`{color: "000000", align: :left, font_size: 10, bold: false}`|Styles for non-header rows. Currently only supports these options (values can be changed though)|
  
<br>
#### `SpreadsheetArchitect.to_csv`
|Option|Type|Default|Notes|
|---|---|---|---|
|**data**|Array| |**Required** 2D Array of data for the non-header row cells.|
|**headers**|Array|`false`|2D Array of data for the header rows cells. Pass false to skip the header row.|



# Change model default method options
```ruby
class Post
  include SpreadsheetArchitect

  def spreadsheet_columns
    [:name, :content]
  end

  SPREADSHEET_OPTIONS = {
    headers: true,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: self.name
  }
end
```

# Change project wide default method options
```ruby
# config/initializers/spreadsheet_architect.rb

SpreadsheetArchitect.module_eval do
  const_set('SPREADSHEET_OPTIONS', {
    headers: true,
    header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    row_style: {background_color: nil, color: 'FFFFFF', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
    sheet_name: 'My Project Export'
  })
end
```


# Credits
Created by Weston Ganger - @westonganger

Heavily influenced by the dead gem `acts_as_xlsx` by @randym but adapted to work for more spreadsheet types and plain ruby models.


<a href="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=VKY8YAWAS5XRQ&lc=CA&item_name=Weston%20Ganger&item_number=spreadsheet_architect&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHostedGuest" target="_blank" title="Donate"><img src="https://www.paypalobjects.com/en_US/i/btn/btn_donate_SM.gif" alt="Donate"/></a>
