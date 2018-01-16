headers = [
  ['Latest Posts'],
  ['Title','Category','Author','Posted on','Earnings']
]

data = @posts.map{|x| [x.title, x.category.name, x.author.name, x.created_at, x.earnings]}

header_style = {}

row_style = {}

column_styles = [
  {columns: 0, styles: {bold: true}},
  {columns: (1..3), styles: {format_code: "$#,##0.00"}},
  {columns: [4,9], include_header: true, styles: {italic: true}}
]

range_styles = [
  {range: "B2:C4", styles: {background_color: "CCCCCC"}},
  {range: {rows: 1, columns: :all}, styles: {bold: true}},
  {range: {rows: (0..5), columns: (1..5)}, styles: {italic: true}},
  {range: {rows: :all, columns: (3..5)}, styles: {color: "FFFFFF"}}
]

borders = [
  {range: "B2:C4"},
  {range: "D6:D7", border_styles: {style: :dashDot, color: "333333"}},
  {range: {rows: (2..11), columns: :all}, border_styles: {edges: [:top,:bottom]}},
  {range: {rows: 3, columns: 4}, border_styles: {edges: [:top,:bottom]}}
]

merges = [
  {range: "A1:C1"},
  {range: {rows: 1, columns: :all}},
  {range: {rows: (1..2), columns: (0..3)}}
]

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

# Save the file_data
File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write file_data
end
