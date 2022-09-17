require "test_helper"

class OdsGeneralTest < ActiveSupport::TestCase

  test "multisheet" do
    test_data = [[1,2,3], [4,5,6], [7,8,9]]

    spreadsheet = Post.to_rodf_spreadsheet
    spreadsheet = CustomPost.to_rodf_spreadsheet({sheet_name: 'Latest Projects'}, spreadsheet)
    spreadsheet = SpreadsheetArchitect.to_rodf_spreadsheet({data: test_data, sheet_name: 'Another Sheet'}, spreadsheet)

    save_file("ods/multi_sheet.ods", spreadsheet.bytes)
  end

  test "kitchen sink" do
    options = {
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Posted on','Posted At','Earnings']
      ],
      data: 50.times.map{|i| [i, "foobar-#{i}", 5.4*i, true, Date.today, Time.now]},
      header_style: {background_color: "000000", color: "FFFFFF", align: :center, font_size: 12, bold: true},
      row_style: {background_color: nil, color: "000000", align: :left, font_size: 12},
      sheet_name: 'Kitchen Sink',
      freeze_headers: true,
      column_types: [
        :string,
        :float,
        :float, 
        :boolean,
        :date,
        :time,
        nil
      ],
    }

    # Using Array Data
    file_data = SpreadsheetArchitect.to_ods(options)

    save_file("ods/kitchen_sink.ods", file_data)
  end

  describe "hyperlinks" do
    test "shows the text and have the correct attributes" do
      url = "https://github.com/westonganger/spreadsheet_architect"

      data = [
        [1,2,3],
        [1, url, "https://github.com/caxlsx/caxlsx"],
      ]

      ss = SpreadsheetArchitect.to_rodf_spreadsheet(data: data, column_types: [:string, :hyperlink, :string])

      doc = parse_ods_spreadsheet(ss)

      cells = doc.xpath("//table:table-cell")
      hyperlinks = doc.xpath("//text:a")

      assert_equal 6, cells.size
      assert_equal 2, hyperlinks.size

      cell = cells[1]
      assert_equal "2", cell.text
      assert_equal "2", cell.at_xpath(".//text:a").attributes["href"].value

      cell = cells[4]
      assert_equal url, cell.text
      assert_equal url, cell.at_xpath(".//text:a").attributes["href"].value

      save_file("ods/hyperlinks.ods", ss)
    end
  end

end
