module SpreadsheetArchitect
  module Utils
    def self.get_cell_data(options, klass)
      if options[:data] && options[:instances]
        raise SpreadsheetArchitect::Exceptions::MultipleDataSourcesError
      elsif options[:data]
        data = options[:data]
      end

      if !options[:data] && options[:headers] == true
        headers = []
        needs_headers = true
      elsif options[:headers].is_a?(Array)
        headers = options[:headers]
      else
        headers = false
      end

      if options[:column_types]
        column_types = options[:column_types]
      else
        column_types = []
        needs_column_types = true
      end

      if !data
        if !options[:instances]
          if is_ar_model?(klass) 
            options[:instances] = klass.where(nil).to_a # triggers the relation call, not sure how this works but it does
          else
            raise SpreadsheetArchitect::Exceptions::NoDataError
          end
        end

        if options[:spreadsheet_columns]
          if [String, Symbol].any?{|x| options[:spreadsheet_columns].is_a?(x)}
            cols_method_name = options[:spreadsheet_columns]

            if klass != SpreadsheetArchitect && !klass.instance_methods.include?(cols_method_name)
              raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError.new(klass, cols_method_name)
            end
          end
        else
          if klass != SpreadsheetArchitect && !klass.instance_methods.include?(:spreadsheet_columns)
            if is_ar_model?(klass)
              the_column_names = klass.column_names
              headers = the_column_names.map{|x| str_titleize(x)} if needs_headers
              columns = the_column_names.map{|x| x.to_sym}
            else
              raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError.new(klass)
            end
          end
        end

        data = []
        options[:instances].each do |instance|
          if columns
            data.push columns.map{|col| col.is_a?(Symbol) ? instance.send(col) : col}
          else
            row_data = []

            if options[:spreadsheet_columns]
              if cols_method_name
                instance_cols = instance.send(cols_method_name)
              else
                instance_cols = options[:spreadsheet_columns].call(instance)
              end
            else
              if klass == SpreadsheetArchitect && !instance.respond_to?(:spreadsheet_columns)
                raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError.new(instance.class)
              else
                instance_cols = instance.spreadsheet_columns
              end
            end

            instance_cols.each_with_index do |x,i|
              if x.is_a?(Array)
                headers.push(x[0].to_s) if needs_headers
                row_data.push(x[1].is_a?(Symbol) ? instance.send(x[1]) : x[1])
                if needs_column_types
                  column_types[i] = x[2]
                end
              else
                headers.push(str_titleize(x.to_s)) if needs_headers
                row_data.push(x.is_a?(Symbol) ? instance.send(x) : x)
              end
            end

            data.push row_data

            needs_headers = false
            needs_column_types = false
          end

        end
      end

      if headers && !headers[0].is_a?(Array)
        headers = [headers]
      end

      if column_types.compact.empty?
        column_types = nil
      end

      return options.merge(headers: headers, data: data, column_types: column_types)
    end

    def self.get_options(options, klass)
      verify_option_types(options)

      if defined?(klass::SPREADSHEET_OPTIONS)
        if klass::SPREADSHEET_OPTIONS.is_a?(Hash)
          options = SpreadsheetArchitect.default_options.merge(
            klass::SPREADSHEET_OPTIONS.merge(options)
          )
        else
          raise SpreadsheetArchitect::Exceptions::OptionTypeError.new("#{klass}::SPREADSHEET_OPTIONS constant")
        end
      else
        options = SpreadsheetArchitect.default_options.merge(options)
      end

      if !options[:headers]
        options[:header_style] = false
      end

      if !options[:sheet_name]
        if klass == SpreadsheetArchitect
          options[:sheet_name] = 'Sheet1'
        else
          options[:sheet_name] = klass.name

          if options[:sheet_name].respond_to?(:pluralize)
            options[:sheet_name] = options[:sheet_name].pluralize
          end
        end
      end

      return options
    end

    def self.convert_styles_to_ods(styles={})
      styles = {} unless styles.is_a?(Hash)
      styles = stringify_keys(styles)

      property_styles = {}

      text_styles = {}
      text_styles['font-weight'] = styles.delete('bold') ? 'bold' : styles.delete('font-weight')
      text_styles['font-size'] = styles.delete('font_size') || styles.delete('font-size')
      text_styles['font-style'] = styles.delete('italic') ? 'italic' : styles.delete('font-style')
      if styles['underline']
        styles.delete('underline')
        text_styles['text-underline-style'] = 'solid'
        text_styles['text-underline-type'] = 'single'
      end
      if styles['align']
        text_styles['align'] = true
      end 
      if styles['color'].respond_to?(:sub) && !styles['color'].empty?
        text_styles['color'] = "##{styles.delete('color').sub('#','')}"
      end
      text_styles.delete_if{|_,v| v.nil?}
      property_styles['text'] = text_styles
      
      cell_styles = {}
      styles['background_color'] ||= styles.delete('background-color')
      if styles['background_color'].respond_to?(:sub) && !styles['background_color'].empty?
        cell_styles['background-color'] = "##{styles.delete('background_color').sub('#','')}"
      end

      cell_styles.delete_if{|_,v| v.nil?}
      property_styles['cell'] = cell_styles

      return property_styles
    end

    private

    def self.is_ar_model?(klass)
      defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
    end

    def self.str_titleize(str)
      str = str.sub(/\A_+/, '')
               .gsub(/[_\.]/,' ')
               .sub(' rescue nil','')
               .gsub(/(\A|\ )\w/){|x| x.upcase}

      return str
    end

    def self.check_option_type(options, option_name, type)
      val = options[option_name]

      if val
        if type.is_a?(Array)
          invalid = type.all?{|t| !val.is_a?(t) }
        elsif !val.is_a?(type)
          invalid = true
        end

        if invalid
          raise SpreadsheetArchitect::Exceptions::OptionTypeError.new(":#{option_name} option")
        end
      end
    end

    def self.verify_option_types(options)
      check_option_type(options, :spreadsheet_columns, [Proc, Symbol, String])
      check_option_type(options, :data, Array)
      check_option_type(options, :instances, Array)
      check_option_type(options, :headers, [TrueClass, Array])
      check_option_type(options, :header_style, Hash)
      check_option_type(options, :row_style, Hash)
      check_option_type(options, :column_styles, Array)
      check_option_type(options, :range_styles, Array)
      check_option_type(options, :conditional_row_styles, Array)
      check_option_type(options, :merges, Array)
      check_option_type(options, :borders, Array)
      check_option_type(options, :column_widths, Array)
    end

    def self.stringify_keys(hash={})
      new_hash = {}
      hash.each do |k,v|
        if v.is_a?(Hash)
          new_hash[k.to_s] = self.stringify_keys(v)
        else
          new_hash[k.to_s] = v
        end
      end
      return new_hash
    end
  end
end
