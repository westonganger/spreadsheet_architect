module SpreadsheetArchitect
  module ClassMethods

    def to_csv(opts={})
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

      CSV.generate do |csv|
        if options[:headers]
          (options[:headers][0].is_a?(Array) ? options[:headers] : [options[:headers]]).each do |header_row|
            csv << options[:headers]
          end
        end
        
        options[:data].each do |row_data|
          csv << row_data
        end
      end
    end

    def to_rodf_spreadsheet
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)

      spreadsheet = ODF::Spreadsheet.new

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
          (options[:headers][0].is_a?(Array) ? options[:headers] : [options[:headers]]).each do |header_row|
            row do
              options[:headers].each_with_index do |header, i|
                cell header, style: :header_style
              end
            end
          end
        end
        options[:data].each_with_index do |row_data, index|
          row do 
            row_data.each_with_index do |y,i|
              cell y, style: :row_style, type: options[:types][i]
            end
          end
        end
      end

      return spreadsheet
    end
    
    def to_ods(opts={})
      return to_rodf_spreadsheet(opts).bytes
    end

    def to_axlsx(which='sheet', opts={})
      opts = SpreadsheetArchitect::Utils.get_cell_data(opts, self)
      options = SpreadsheetArchitect::Utils.get_options(opts, self)
    
      header_style = {}
      if options[:header_style]
        header_style[:fg_color] = options[:header_style].delete(:color)
        header_style[:bg_color] = options[:header_style].delete(:background_color)
        if header_style[:align]
          header_style[:alignment] = {}
          header_style[:alignment][:horizontal] = options[:header_style].delete(:align)
        end
        header_style[:b] = options[:header_style].delete(:bold)
        header_style[:sz] = options[:header_style].delete(:font_size)
        header_style[:i] = options[:header_style].delete(:italic)
        if options[:header_style][:underline]
          header_style[:u] = options[:header_style].delete(:underline)
        end
        header_style.delete_if{|x| x.nil?}

        header_style = options[:header_style].merge(header_style)
      end


      row_style = {}
      if options[:row_style]
        row_style[:fg_color] = options[:row_style].delete(:color)
        row_style[:bg_color] = options[:row_style].delete(:background_color)
        if row_style[:align]
          row_style[:alignment] = {}
          row_style[:alignment][:horizontal] = options[:row_style][:align]
        end
        row_style[:b] = options[:row_style].delete(:bold)
        row_style[:sz] = options[:row_style].delete(:font_size)
        row_style[:i] = options[:row_style].delete(:italic)
        if options[:row_style][:underline]
          row_style[:u] = options[:row_style].delete(:underline)
        end
        row_style.delete_if{|x| x.nil?}

        row_style = options[:row_style].merge(row_style)
      end
      
      package = Axlsx::Package.new

      return package if options[:data].empty?

      the_sheet = package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|

        if options[:headers]
          options[:headers] = [options[:headers]] unless options[:headers][0].is_a?(Array)

          options[:headers].each do |header_row|
            sheet.add_row header_row, style: package.workbook.styles.add_style(header_style)
          end
        end
        
        options[:data].each do |row_data|
          sheet.add_row row_data, style: package.workbook.styles.add_style(row_style), types: options[:types]
        end

        options[:types].each_with_index do |type, index|
          if [:date, :time].include?(type)
            if type == :date
              format_code = 'm/d/yyyy'
            else
              format_code = 'm/d/yyyy h:mm AM/PM'
            end

            sheet.col_style(index, {format_code: format_code}, row_offset: (options[:headers] ? options[:headers].count : 0))
          end
        end

        if options[:column_styles]
          options[:column_styles].each do |x|
            if x[:include_header] || !options[:headers]
              start_row = 0
            else
              start_row = options[:headers].count
            end

            if x[:columns].is_a?(Array) || x[:columns].is_a?(Range) 
              x[:columns].each do |col|
                sheet.col_style(col, x[:styles], row_offset: start_row)
              end
            elsif x[:columns].is_a?(Integer)
              sheet.col_style(x[:columns], x[:styles], row_offset: start_row)
            end
          end
        end

        col_names = Array('A'..'ZZZ')
        row_count = options[:data].count
        row_count += options[:headers].count if options[:headers]

        if options[:borders]
          options[:borders].each do |x|
            range = ''
            if x[:columns]
              start_row = (x[:include_headers] && options[:headers]) ? options[:headers].count : 0

              if x[:columns].is_a?(Array) || x[:columns].is_a?(Range) 
                x[:columns].each do |col|
                  range = "#{col}#{start_row}:#{col}#{row_count}"
                end
              elsif x[:columns].is_a?(Integer)
                col_letter = col_names[x[:columns]]
                range = "#{col_letter}#{start_row}:#{col_letter}#{row_count}"
              end
            end

            if x[:rows]

            end

            if x[:range]

            end

            sheet.add_border range, x[:border_styles]
          end
        end

        if options[:range_styles]
          options[:range_styles].each do |range, styles|
            sheet.add_style range, styles
          end
        end
      end

      if which.to_sym == :sheet
        return the_sheet
      else
        return package
      end
    end

    def to_xlsx(opts={})
      return to_axlsx('package', opts).to_stream.read
    end

  end
end
