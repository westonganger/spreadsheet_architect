require "test_helper"

class KitchenSinkTest < ActiveSupport::TestCase

  def setup
    @options = {
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Posted on','Posted At','Earnings']
      ],
      data: 50.times.map{|i| [i, "foobar-#{i}", 5.4*i, true, Date.today, Time.now]},
      header_style: {background_color: "000000", color: "FFFFFF", align: :center, font_size: 12, bold: true},
      row_style: {background_color: nil, color: "000000", align: :left, font_size: 12},
      sheet_name: 'Kitchen Sink',
      freeze_headers: true,
    }
  end

  def teardown
  end

  def test_ods
    opts = @options.merge({
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Boolean','Posted on','Posted At']
      ],
      data: 50.times.map{|i| [i, "foobar-#{i}", (5.4*i), true, Date.today, Time.now]},
      column_types: [
        :string,
        :float,
        :float, 
        :boolean,
        :date,
        :time,
        nil
      ],
    })

    # Using Array Data
    file_data = SpreadsheetArchitect.to_ods(opts)

    File.open(TMP_PATH.join("kitchen_sink.ods"),'w+b') do |f|
      f.write file_data
    end
  end

end
