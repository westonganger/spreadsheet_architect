require 'axlsx'
require 'axlsx_styler'

require 'spreadsheet_architect/monkey_patches/axlsx_column_width'

module SpreadsheetArchitect
  module ClassMethods
    def to_xlsx(opts={})
      return to_axlsx_package(opts).to_stream.read
    end

    def to_axlsx_package(opts={}, package=nil)
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)
    
      header_style = SpreadsheetArchitect::Utils.convert_styles_to_axlsx(options[:header_style])
      row_style = SpreadsheetArchitect::Utils.convert_styles_to_axlsx(options[:row_style])

      if package.nil?
        package = Axlsx::Package.new
      end

      return package if !options[:headers] && options[:data].empty?

      package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|
        if options[:headers]
          header_style_index = package.workbook.styles.add_style(header_style)

          options[:headers].each do |header_row|
            sheet.add_row header_row, style: header_style_index
          end
        end

        if options[:data].empty?
          break
        end

        row_style_index = package.workbook.styles.add_style(row_style)

        max_row_length = options[:data].max_by{|x| x.length}.length

        options[:data].each do |row_data|
          missing = max_row_length - row_data.count
          if missing > 0
            missing.times do
              row_data.push nil
            end
          end

          types = []
          row_data.each_with_index do |x,i|
            if options[:column_types]
              types[i] = options[:column_types][i]
            end

            types[i] ||= SpreadsheetArchitect::Utils.get_type(x)
          end

          sheet.add_row row_data, style: row_style_index, types: types
        end

        options[:data].first.each_with_index do |x,i|
          types = []

          if options[:column_types]
            types[i] = options[:column_types][i]
          end

          types[i] ||= SpreadsheetArchitect::Utils.get_type(x)

          if [:date, :time].include?(types[i])
            if types[i] == :date
              format_code = 'm/d/yyyy'
            else
              format_code = 'yyyy/m/d h:mm AM/PM'
            end

            sheet.col_style(i, package.workbook.styles.add_style(format_code: format_code), row_offset: (options[:headers] ? options[:headers].count : 0))
          end
        end

        if options[:column_widths]
          sheet.column_widths(*options[:column_widths])
        end

        if options[:borders] || options[:column_styles]
          col_names = max_row_length > 675 ? Array('A'..'ZZZ') : Array('A'..'ZZ')
        end

        if options[:borders]
          row_count = options[:data].count
          row_count += options[:headers].count if options[:headers]

          options[:borders].each do |x|
            if x[:columns]
              start_row = 1
              start_row += options[:headers].count if !x[:include_header] && options[:headers]

              if x[:columns].is_a?(Array)
                x[:columns].each do |col|
                  if col.is_a?(Integer)
                    col_names[col]
                  end
                  sheet.add_border "#{col}#{start_row}:#{col}#{row_count}", x[:border_styles]
                end
              elsif x[:columns].is_a?(Range)
                if x[:columns].first.is_a?(Integer)
                  range = "#{col_names[x[:columns].first]}#{start_row}:#{col_names[x[:columns].last]}#{row_count}"
                else
                  range = "#{x[:columns].first}#{start_row}:#{x[:columns].last}#{row_count}"
                end
                sheet.add_border range, x[:border_styles]
              elsif x[:columns].is_a?(Integer)
                col_letter = col_names[x[:columns]]
                sheet.add_border "#{col_letter}#{start_row}:#{col_letter}#{row_count}", x[:border_styles]
              elsif x[:columns].is_a?(String)
                sheet.add_border "#{x[:columns]}#{start_row}:#{x[:columns]}#{row_count}", x[:border_styles]
              end
            end

            if x[:rows]
              start_col = x[:start_column] ? x[:start_column] : 'A'
              if x[:end_column]
                end_col = x[:end_column] > col_names[max_row_length-1] ? col_names[max_row_length-1] : x[:end_column]
              else
                end_col = col_names[max_row_length-1]
              end

              if x[:rows].is_a?(Array)
                x[:rows].each do |row|
                  sheet.add_border "#{start_col}#{row}:#{end_col}#{row}", x[:border_styles]
                end
              elsif x[:rows].is_a?(Range)
                sheet.add_border "#{start_col}#{x[:rows].first}:#{end_col}#{x[:rows].last}", x[:border_styles]
              elsif x[:rows].is_a?(Integer)
                sheet.add_border "#{start_col}#{x[:rows]}:#{end_col}#{x[:rows]}", x[:border_styles]
              end
            end

            if x[:range]
              sheet.add_border x[:range], x[:border_styles]
            end
          end
        end

        if options[:column_styles]
          options[:column_styles].each do |x|
            start_row = !x[:include_header] && options[:headers] ? options[:headers].count : 0

            package.workbook.styles do |s|
              style = s.add_style row_style.merge(SpreadsheetArchitect::Utils.convert_styles_to_axlsx(x[:styles]))
              if x[:columns].is_a?(Array) || x[:columns].is_a?(Range) 
                x[:columns].each do |col|
                  if col.is_a?(String)
                    col = col_names.index(col)
                  end

                  sheet.col_style(col, style, row_offset: start_row)
                end
              elsif x[:columns].is_a?(Integer)
                sheet.col_style(x[:columns], style, row_offset: start_row)
              end
            end
          end
        end

        if options[:range_styles]
          options[:range_styles].each do |x|
            styles = SpreadsheetArchitect::Utils.convert_styles_to_axlsx(x[:styles])

            sheet.add_style x[:range], styles
          end
        end

        if options[:merges]
          options[:merges].each do |x|
            sheet.merge_cells x[:range]
          end
        end
      end

      return package
    end
  end
end
