module SpreadsheetArchitect
  module ClassMethods
    def to_csv(opts={})
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

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
    
    def to_ods(opts={})
      return to_rodf_spreadsheet(opts).bytes
    end

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

      return package if options[:data].empty?

      package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|
        if options[:headers]
          header_style_index = package.workbook.styles.add_style(header_style)

          options[:headers].each do |header_row|
            sheet.add_row header_row, style: header_style_index
          end
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
          if options[:column_types]
            types[i] = options[:column_types][i]
          end

          types[i] ||= SpreadsheetArchitect::Utils.get_type(x)

          if [:date, :time].include?(types[i])
            if types[i] == :date
              format_code = 'm/d/yyyy'
            else
              format_code = 'm/d/yyyy h:mm AM/PM'
            end

            sheet.col_style(i, package.workbook.styles.add_style(format_code: format_code), row_offset: (options[:headers] ? options[:headers].count : 0))
          end
        end

        if options[:column_widths]
          sheet.column_widths options[:column_widths]
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
              start_col = x[:start_offset] ? col_names[x[:start_offset].to_i] : 'A'

              end_index = max_row_length - 1
              end_index -= x[:end_offset].to_i if x[:end_offset]
              end_col = col_names[end_index]

              if x[:rows].is_a?(Array)
                x[:rows].each do |row|
                  sheet.add_border "#{start_col}#{row}:#{end_col}#{row}", x[:border_styles]
                end
              elsif x[:rows].is_a?(Range)
                sheet.add_border "#{start_col}#{x[:rows].first}:#{end_col}#{x[:rows].last}", x[:border_styles]
              elsif x[:rows].is_a?(Integer)
                col = col_names[x[:columns]]
                sheet.add_border "#{start_col}#{col}:#{end_col}#{col}", x[:border_styles]
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
            header_count = 1
            header_count += x[:headers].count if x[:headers]

            #range_numbers = x[:range].scan(/\d+/).map{|num| num.to_i}

            #if range_numbers.first < header_count && range_numbers.last < header_count
            #  styles = header_style.merge(SpreadsheetArchitect::Utils.convert_styles_to_axlsx(x[:styles]))
            #else
            #  styles = row_style.merge(SpreadsheetArchitect::Utils.convert_styles_to_axlsx(x[:styles]))
            #end

            # with new axlsx_styler patch we dont need to provide a workaround for compounding styles
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

    def to_rodf_spreadsheet(opts={}, spreadsheet=nil)
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

      if !spreadsheet
        spreadsheet = RODF::Spreadsheet.new
      end

      spreadsheet.office_style :header_style, family: :cell do
        if options[:header_style]
          unless opts[:header_style] && opts[:header_style][:bold] == false #uses opts, temporary
            property :text, 'font-weight' => :bold
          end
          if options[:header_style][:align]
            property :text, align: options[:header_style][:align]
          end
          if options[:header_style][:size]
            property :text, 'font-size' => options[:header_style][:size]
          end
          if options[:header_style][:color] && opts[:header_style] && opts[:header_style][:color] #temporary
            property :text, color: "##{options[:header_style][:color]}"
          end
        end
      end
      spreadsheet.office_style :row_style, family: :cell do
        if options[:row_style]
          if options[:row_style][:bold]
            property :text, 'font-weight' => :bold
          end
          if options[:row_style][:align]
            property :text, align: options[:row_style][:align]
          end
          if options[:row_style][:size]
            property :text, 'font-size' => options[:row_style][:size]
          end
          if opts[:row_style] && opts[:row_style][:color] #uses opts, temporary
            property :text, color: "##{options[:row_style][:color]}"
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
