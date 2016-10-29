headers = [
  ['Latest Posts'],
  ['Title','Category','Author','Posted on','Earnings']
]

data = @posts.map{|x| [x.title, x.category.name, x.author.name, x.created_at, x.earnings]}

header_style = {}

row_style = {}

column_styles = [
  {columns: 0, styles: {bold: true}},
  {columns: (1..3), styles: {number_format: "$#,##0.00"}},
  {columns: [4,9], include_header: true, styles: {italic: true}}
]

range_styles = [
  {range: "B2:C4", styles: {background_color: "CCCCCC"}}
]

borders = [
  {range: "B2:C4"},
  {range: "D6:D7", border_styles: {style: :dashdot, color: "333333"}},
  {rows: (2..11), border_styles: {edges: [:top,:bottom]}},
  {rows: [1,3,5], start_column: 'B', end_column: 'F'},
  {rows: 1},
  {columns: 0, border_styles: {edges: [:right], style: :thick}},
  {columns: (1..2)},
  {columns: [4,6,8]}

]

merges = [
  {range: "A1:C1"}
]

# Using Array Data
file_data = SpreadsheetArchitect.to_xlsx({
  headers: headers,
  data: data,
  header_style: header_style,
  row_style: row_style,
  column_styles: column_styles,
  range_styles: range_styles,
  borders: borders,
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
