require 'rodf'

module SpreadsheetArchitect
  module ClassMethods
    def to_ods(opts={})
      return to_rodf_spreadsheet(opts).bytes
    end

    def to_rodf_spreadsheet(opts={}, spreadsheet=nil)
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

      if !spreadsheet
        spreadsheet = RODF::Spreadsheet.new
      end

      spreadsheet.office_style :header_style, family: :cell do
        if options[:header_style]
          SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:header_style]).each do |k,v|
            property v['property'], k => v['value']
          end
        end
      end
      spreadsheet.office_style :row_style, family: :cell do
        if options[:row_style]
          SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:row_style]).each do |k,v|
            property v['property'], k => v['value']
          end
        end
      end

      spreadsheet.table options[:sheet_name] do 
        if options[:headers]
          options[:headers].each do |header_row|
            row do
              header_row.each_with_index do |header, i|
                cell header, style: :header_style
              end
            end
          end
        end

        options[:data].each_with_index do |row_data, index|
          row do 
            row_data.each_with_index do |y,i|
              cell y, style: :row_style, type: (options[:types][i] if options[:types])
            end
          end
        end
      end

      return spreadsheet
    end

  end
end
