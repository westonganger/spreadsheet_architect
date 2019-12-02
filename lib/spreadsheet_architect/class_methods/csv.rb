require 'csv'

module SpreadsheetArchitect
  module ClassMethods

    def to_csv(opts={})
      opts = SpreadsheetArchitect::Utils.get_options(opts, self)
      options = SpreadsheetArchitect::Utils.get_cell_data(opts, self)

      CSV.generate do |csv|
        if options[:headers]
          options[:headers].each do |header_row|
            csv << header_row
          end
        end
        
        options[:data].each do |row_data|
          csv << row_data
        end
      end
    end

  end
end
