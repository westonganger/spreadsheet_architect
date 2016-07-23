headers = [
  ['Latest Posts'],
  ['Title','Category','Author','Posted on','Earnings']
]

data = @posts.map{|x| [x.title, x.category.name, x.author.name, x.created_at, x.earnings]}

header_style = {}

row_style = {}

# zero based column indexes
column_styles = [
  {},
  {},
  {},
  {},
  {}
]

# uses 1 based rows just like the spreadsheet grid
range_styles = [
  {},
  {},
  {},
  {},
  {}
]

borders = [
  {},
  {},
  {},
  {},
  {}
]

merges = [
  {},
  {},
  {},
  {},
  {}
]

# Using Array Data
file_data = SpreadsheetArchitect.to_xlsx({
  headers: headers,
  data: data,
  header_style: header_style,
  row_style: row_style,
  column_styles: column_styles,
  range_styles: range_styles,
  #borders: borders, # commented out because they currently cause problems with styles
  merges: merges
})

# OR

# With Rails ActiveRecord Model
file_data = Post.order(created_at: :desc).limit(100).to_xlsx({
  headers: headers,
  data: data,
  header_style: header_style,
  row_style: row_style,
  column_styles: column_styles,
  range_styles: range_styles,
  borders: borders,
  merges: merges
})

# OR with Plain Ruby Class
file_data = Post.to_xlsx({
  instances: @posts,
  headers: headers,
  data: data,
  header_style: header_style,
  row_style: row_style,
  column_styles: column_styles,
  range_styles: range_styles,
  borders: borders,
  merges: merges
})


# Save the file_data
File.open('path/to/file.xlsx') do |f|
  f.write file_data
end
