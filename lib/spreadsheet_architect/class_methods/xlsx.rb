require 'axlsx'
require 'axlsx_styler'

require 'spreadsheet_architect/axlsx_string_width_patch'

module SpreadsheetArchitect
  module ClassMethods

    def to_xlsx(opts={})
      return to_axlsx_package(opts).to_stream.read
    end

    def to_axlsx_package(opts={}, package=nil)
      opts = SpreadsheetArchitect::Utils.get_options(opts, self)
      options = SpreadsheetArchitect::Utils.get_cell_data(opts, self)

      header_style = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(options[:header_style])
      row_style = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(options[:row_style])

      if package.nil?
        package = Axlsx::Package.new
      end

      row_index = -1

      package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|
        max_row_length = options[:data].empty? ? 0 : options[:data].max_by{|x| x.length}.length

        if options[:headers]
          header_style_index = package.workbook.styles.add_style(header_style)

          options[:headers].each do |header_row|
            row_index += 1

            missing = max_row_length - header_row.count
            if missing > 0
              missing.times do
                header_row.push(nil)
              end
            end

            sheet.add_row header_row, style: header_style_index

            if options[:conditional_row_styles]
              conditional_styles_for_row = SpreadsheetArchitect::Utils::XLSX.conditional_styles_for_row(options[:conditional_row_styles], row_index, header_row)

              unless conditional_styles_for_row.empty?
                sheet.add_style(
                  "#{SpreadsheetArchitect::Utils::XLSX::COL_NAMES.first}#{row_index+1}:#{SpreadsheetArchitect::Utils::XLSX::COL_NAMES[max_row_length-1]}#{row_index+1}",
                  SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(conditional_styles_for_row)
                )
              end
            end
          end
        end

        if options[:data].empty?
          break
        end

        row_style_index = package.workbook.styles.add_style(row_style)

        default_date_style_index = nil
        default_time_style_index = nil

        options[:data].each do |row_data|
          row_index += 1

          missing = max_row_length - row_data.count
          if missing > 0
            missing.times do
              row_data.push(nil)
            end
          end

          types = []
          styles = []
          row_data.each_with_index do |x,i|
            if (x.respond_to?(:empty) ? x.empty? : x.nil?)
              types[i] = nil
              styles[i] = row_style_index
            else
              if options[:column_types]
                types[i] = options[:column_types][i]
              end

              types[i] ||= SpreadsheetArchitect::Utils::XLSX.get_type(x)

              if [:date, :time].include?(types[i])
                if types[i] == :date
                  default_date_style_index ||= package.workbook.styles.add_style(row_style.merge({format_code: 'yyyy-mm-dd'}))
                  styles[i] = default_date_style_index
                else
                  default_time_style_index ||= package.workbook.styles.add_style(row_style.merge({format_code: 'yyyy-mm-dd h:mm AM/PM'}))
                  styles[i] = default_time_style_index
                end
              else
                styles[i] = row_style_index
              end
            end
          end

          sheet.add_row row_data, style: styles, types: types, escape_formulas: options[:escape_formulas]

          if options[:conditional_row_styles]
            options[:conditional_row_styles] = SpreadsheetArchitect::Utils.hash_array_symbolize_keys(options[:conditional_row_styles])

            conditional_styles_for_row = SpreadsheetArchitect::Utils::XLSX.conditional_styles_for_row(options[:conditional_row_styles], row_index, row_data)

            unless conditional_styles_for_row.empty?
              sheet.add_style(
                "#{SpreadsheetArchitect::Utils::XLSX::COL_NAMES.first}#{row_index+1}:#{SpreadsheetArchitect::Utils::XLSX::COL_NAMES[max_row_length-1]}#{row_index+1}",
                SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(conditional_styles_for_row)
              )
            end
          end
        end

        if options[:column_widths]
          sheet.column_widths(*options[:column_widths])
        end

        if options[:borders] || options[:column_styles] || options[:range_styles] || options[:merges]
          num_rows = options[:data].count + (options[:headers] ? options[:headers].count : 0)
        end

        if options[:borders]
          options[:borders] = SpreadsheetArchitect::Utils.hash_array_symbolize_keys(options[:borders])

          options[:borders].each do |x|
            if x[:range].is_a?(Hash)
              x[:range] = SpreadsheetArchitect::Utils::XLSX.range_hash_to_str(x[:range], max_row_length, num_rows, use_zero_based_row_index: options[:use_zero_based_row_index])
            else
              SpreadsheetArchitect::Utils::XLSX.verify_range(x[:range], num_rows)
            end

            sheet.add_border x[:range], (x[:border_styles] || x[:styles])
          end
        end

        if options[:column_styles]
          options[:column_styles] = SpreadsheetArchitect::Utils.hash_array_symbolize_keys(options[:column_styles])

          options[:column_styles].each do |x|
            start_row = (options[:headers] ? options[:headers].count : 0) + 1

            x[:styles] = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(x[:styles])

            add_column_style = ->(col){
              SpreadsheetArchitect::Utils::XLSX.verify_column(col, max_row_length)

              range_str = SpreadsheetArchitect::Utils::XLSX.range_hash_to_str({rows: (start_row..num_rows), columns: col}, max_row_length, num_rows)
              sheet.add_style range_str, x[:styles]

              if x[:include_header] && start_row > 1
                range_str = SpreadsheetArchitect::Utils::XLSX.range_hash_to_str({rows: (1..start_row-1), columns: col}, max_row_length, num_rows)
                sheet.add_style(range_str, x[:styles])
              end
            }

            case x[:columns]
            when Array, Range
              x[:columns].each do |col|
                add_column_style.call(col)
              end
            when Integer, String
              add_column_style.call(x[:columns])
            else
              SpreadsheetArchitect::Utils::XLSX.verify_column(x[:columns], max_row_length)
            end
          end
        end

        if options[:range_styles]
          options[:range_styles] = SpreadsheetArchitect::Utils.hash_array_symbolize_keys(options[:range_styles])

          options[:range_styles].each do |x|
            styles = SpreadsheetArchitect::Utils::XLSX.convert_styles_to_axlsx(x[:styles])

            if x[:range].is_a?(Hash)
              x[:range] = SpreadsheetArchitect::Utils::XLSX.range_hash_to_str(x[:range], max_row_length, num_rows, use_zero_based_row_index: options[:use_zero_based_row_index])
            else
              SpreadsheetArchitect::Utils::XLSX.verify_range(x[:range], num_rows)
            end

            sheet.add_style x[:range], styles
          end
        end

        if options[:merges]
          options[:merges] = SpreadsheetArchitect::Utils.hash_array_symbolize_keys(options[:merges])

          options[:merges].each do |x|
            if x[:range].is_a?(Hash)
              x[:range] = SpreadsheetArchitect::Utils::XLSX.range_hash_to_str(x[:range], max_row_length, num_rows, use_zero_based_row_index: options[:use_zero_based_row_index])
            else
              SpreadsheetArchitect::Utils::XLSX.verify_range(x[:range], num_rows)
            end

            sheet.merge_cells x[:range]
          end
        end

        if options[:freeze_headers]
          sheet.sheet_view.pane do |pane|
            pane.state = :frozen
            pane.y_split = options[:headers].count
          end

        elsif options[:freeze]
          case options[:freeze][:type].to_s
          when "split_panes"
            options[:freeze][:state] == "split"
          when "frozen", "freeze"
            options[:freeze][:state] == "frozen"
          end

          if options[:freeze][:rows]
            options[:freeze][:row] ||= options[:freeze][:rows]
          end

          if options[:freeze][:columns]
            options[:freeze][:column] ||= options[:freeze][:columns]
          end

          if options[:freeze][:row] == :all
            options[:freeze][:row] = nil
          elsif options[:freeze][:row].is_a?(Range)
            options[:freeze][:row] = options[:freeze][:row].last
          end

          if options[:freeze][:column] == :all
            options[:freeze][:column] = nil
          elsif options[:freeze][:column].is_a?(Range)
            options[:freeze][:column] = options[:freeze][:column].last
          end

          if !options[:freeze][:row] && !options[:freeze][:column]
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Missing required :row or :column value in the :freeze option hash")
          elsif options[:freeze][:row] && !options[:freeze][:row].is_a?(Integer)
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Invalid :row value provided for in :freeze option hash, must be an Integer")
          elsif options[:freeze][:column] && !options[:freeze][:column].is_a?(Integer)
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Invalid :column value provided for in :freeze option hash, must be an Integer")
          end

          sheet.sheet_view.pane do |pane|
            pane.state = (options[:freeze][:state] || :frozen).to_sym ### Other options are :split and :frozen_split

            if options[:freeze][:active_pane]
              pane.active_pane = options[:freeze][:active_pane].to_sym
            end

            if options[:freeze][:top_left_cell]
              pane.top_left_cell = options[:freeze][:top_left_cell]
            end

            if options[:freeze][:row]
              if options[:use_zero_based_row_index]
                options[:freeze][:row] += 1
              end

              pane.y_split = options[:freeze][:row]
            end

            if options[:freeze][:column]
              pane.x_split = options[:freeze][:column]
            end
          end
        end
      end

      return package
    end

  end
end
