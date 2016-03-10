# Spreadsheet Architect
<a href="https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=VKY8YAWAS5XRQ&lc=CA&item_name=Weston%20Ganger&item_number=spreadsheet_architect&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHostedGuest" target="_blank" title="Donate"><img src="https://www.paypalobjects.com/en_US/i/btn/btn_donate_SM.gif" alt="Donate"/></a>

Spreadsheet Architect lets you turn any activerecord relation or plain ruby class object into a XLSX, ODS, or CSV spreadsheets. Generates columns from model activerecord column_names or from an array of ruby methods.

Spreadsheet Architect adds the following methods to your class:
```ruby
# Plain Ruby
Post.to_xlsx(data: posts_array)
Post.to_ods(data: posts_array)
Post.to_csv(data: posts_array)

# Rails
Post.order(name: :asc).where(published: true).to_xlsx
Post.order(name: :asc).where(published: true).to_ods
Post.order(name: :asc).where(published: true).to_csv
```

## Note: Major bug fix in 1.2.0
v.1.1.0 had a major bug in it where it would use only the first records data for non-symbol methods. Please update to 1.2.0 from 1.1.0 as soon as possible.

## Note: Breaking Changes in 1.1.0
The `spreadsheet_columns` method has been moved from the class to the instance. So now you can use string interpolation in your values. Please re-read the Model section below to see the changes. The side effect of this is if you are using the spreadsheet_columns option directly on the .to_* methods.


# Install
```ruby
gem install spreadsheet_architect
```


# Setup

### Model
```ruby
class Post < ActiveRecord::Base #activerecord not required
  include SpreadsheetArchitect

  belongs_to :author
  belongs_to :category
  has_many :tags

  #optional for activerecord classes, defaults to the models column_names
  def spreadsheet_columns
    #[[Label, Method/Statement to Call on each Instance]....]
    [
      ['Title', :title],
      ['Content', content],
      ['Author', (author.name rescue nil)],
      ['Published?', (published ? 'Yes' : 'No')],
      ['Category/Tags', "#{category.name} - #{tags.collect(&:name).join(', ')}"]
    ]

    # OR just humanize the method to use as the label ex. "Title", "Content", "Author Name", "Published"
    [:title, content, (author.name rescue nil), :published]

    # OR a Combination of Both
    [:title, :content, ['Author',(author.name rescue nil)], :published, ]
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
      format.ods { render ods: Post.to_odf(data: @posts) }
      format.csv{ render csv: @posts.to_csv(headers: false), file_name: 'articles' }
    end
  end
end
```

### Method 2: Save to a file manually
```ruby
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
  f.write{ Post.to_xlsx(data: posts_array) }
end
```


# Method Options

### to_xlsx, to_ods, to_csv
**data** - *Array* - Mainly for Plain Ruby objects pass in an array of instances. Optional for ActiveRecord relations, you can just chain the method to the end of your relation. If Plain Ruby object it defaults to the instances `to_a` method.

**headers** - *Boolean* - Default: true - Pass in false if you do not want a header row.

**spreadsheet_columns** - *Array* - Use this to override the models spreadsheet_columns/column_names method for one time. Must use symbols that correspond to instance methods of the object. Ex: `[:name, :title, :address]` or `[['Name',:name],['Post Title', :title],['Address', :address]]`

### to_xlsx
**sheet_name** - *String*

**header_style** - *Hash* - Default: `{background_color: "AAAAAA", color: "FFFFFF", align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`
  
**row_style** - Hash - Default: `{background_color: nil, color: "FFFFFF", align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false}`

### to_ods
**sheet_name** - *String*

**header_style** - *Hash* - Default: {color: "000000", align: :center, font_size: 10, bold: true} - Note: Currently only supports these options

**row_style** - *Hash* - Default: {color: "000000", align: :left, font_size: 10, bold: false} - Note: Currently only supports these options

### to_csv
Only the generic options


# Credits
Created by Weston Ganger - @westonganger

Heavily influenced by the dead gem `acts_as_xlsx` by @randym but adapted to work for more spreadsheet types and plain ruby models.
