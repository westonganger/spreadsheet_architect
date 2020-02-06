require "test_helper"

class XlsxFreezeTest < ActiveSupport::TestCase

  def setup
    @options = {
      headers: [
        ['Latest Posts'],
        ['Title','Category','Author','Posted on','Posted At','Earnings']
      ],
      data: 50.times.map{|i| [i, "foobar-#{i}", 5.4*i, true, Date.today, Time.now]},
    }
  end

  def teardown
  end

  def test_1
    opts = @options.merge({
      freeze: {rows: 1, columns: 1},
    })

    # Using Array Data
    file_data = SpreadsheetArchitect.to_xlsx(opts)

    File.open(VERSIONED_BASE_PATH.join("freeze_test_1.xlsx"),'w+b') do |f|
      f.write file_data
    end
  end

  def test_2
    opts = @options.merge({
      freeze: {rows: (2..4), columns: (2..4)},
    })

    # Using Array Data
    file_data = SpreadsheetArchitect.to_xlsx(opts)

    File.open(VERSIONED_BASE_PATH.join("freeze_test_2.xlsx"),'w+b') do |f|
      f.write file_data
    end
  end

end
