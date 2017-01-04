# XLSX Multi Spreadsheet

package = Post.to_axlsx_package
package = Project.limit(100).to_axlsx_package({sheet_name: 'Latest Projects'}, package)
package = SpreadsheetArchitect.to_axlsx_package({data: test_data, sheet_name: 'Latest Projects'}, package)

File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write package.to_stream.read
end



# ODS Multi Spreadsheet

spreadsheet = Post.to_rodf_spreadsheet
spreadsheet = Project.limit(100).to_rodf_spreadsheet({sheet_name: 'Latest Projects'}, spreadsheet)
spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({data: test_data, sheet_name: 'Latest Projects'}, spreadsheet)


File.open('path/to/file.xlsx', 'w+b') do |f|
  f.write spreadsheet.bytes
end
