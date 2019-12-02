require 'rodf'

module SpreadsheetArchitect
  module ClassMethods

    def to_ods(opts={})
      return to_rodf_spreadsheet(opts).bytes
    end

    def to_rodf_spreadsheet(opts={}, spreadsheet=nil)
      opts = SpreadsheetArchitect::Utils.get_options(opts, self)
      options = SpreadsheetArchitect::Utils.get_cell_data(opts, self)

      if options[:column_types] && !(options[:column_types].compact.collect(&:to_sym) - SpreadsheetArchitect::ODS_COLUMN_TYPES).empty?
        raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Invalid column type. Valid ODS values are #{SpreadsheetArchitect::ODS_COLUMN_TYPES}")
      end

      if !spreadsheet
        spreadsheet = RODF::Spreadsheet.new
      end

      spreadsheet.office_style :header_style, family: :cell do
        if options[:header_style]
          SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:header_style]).each do |prop, styles|
            styles.each do |k,v|
              property prop.to_sym, k => v
            end
          end
        end
      end

      spreadsheet.office_style :row_style, family: :cell do
        if options[:row_style]
          SpreadsheetArchitect::Utils.convert_styles_to_ods(options[:row_style]).each do |prop, styles|
            styles.each do |k,v|
              property prop.to_sym, k => v
            end
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
            row_data.each_with_index do |val, i|
              if options[:column_types]
                type = options[:column_types][i]
              end

              if (type && [:date, :time].include?(type)) || val.respond_to?(:strftime)
                type = :string
                val = val.to_s 
              end

              cell val, style: :row_style, type: type
            end
          end
        end
      end

      return spreadsheet
    end

  end
end
