module SpreadsheetArchitect
  module Utils
    def self.str_humanize(str, capitalize = true)
      str = str.sub(/\A_+/, '').gsub(/[_\.]/,' ').sub(' rescue nil','')
      if capitalize
        str = str.gsub(/(\A|\ )\w/){|x| x.upcase}
      end
      return str
    end

    def self.get_cell_data(options, klass)
      if options[:data] && options[:instances]
        raise SpreadsheetArchitect::Exceptions::MultipleDataSourcesError
      end

      if options[:data]
        data = options[:data]
      end

      if options[:headers].nil?
        headers = klass::SPREADSHEET_OPTIONS[:headers] if defined?(klass::SPREADSHEET_OPTIONS)
        headers ||= SpreadsheetArchitect.default_options[:headers]
        if headers == true
          headers = []
          needs_headers = true
        end
      elsif options[:headers].is_a?(Array)
        headers = options[:headers]
      else
        headers = false
      end

      if !data
        if !options[:instances] && self.is_ar_model?(klass) 
          options[:instances] = klass.where(nil).to_a # triggers the relation call, not sure how this works but it does
        end

        if !options[:instances]
          raise SpreadsheetArchitect::Exceptions::NoDataError
        end

        has_custom_columns = !!options[:spreadsheet_columns]

        if has_custom_columns
          columns = []

          options[:spreadsheet_columns].each_with_index do |x,i|
            if x.is_a?(Array)
              headers.push(x[0].to_s) if needs_headers
              columns.push x[1]
            else
              headers.push(str_humanize(x.to_s)) if needs_headers
              columns.push x
            end
          end
        else
          if klass != SpreadsheetArchitect && !klass.instance_methods.include?(:spreadsheet_columns)
            if self.is_ar_model?(klass)
              the_column_names = klass.column_names
              headers = the_column_names.map{|x| str_humanize(x)} if needs_headers
              columns = the_column_names.map{|x| x.to_sym}
            else
              raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError.new(klass)
            end
          end
        end

        data = []
        options[:instances].each do |instance|
          if columns
            data.push columns.map{|col| col.is_a?(Symbol) ? instance.instance_eval(col.to_s) : col}
          else
            row_data = []

            if klass == SpreadsheetArchitect && !instance.respond_to?(:spreadsheet_columns)
              raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError.new(instance.class)
            else
              instance_cols = instance.spreadsheet_columns
            end

            instance_cols.each do |x|
              if x.is_a?(Array)
                headers.push(x[0].to_s) if needs_headers
                row_data.push(x[1].is_a?(Symbol) ? instance.instance_eval(x[1].to_s) : x[1])
              else
                headers.push(str_humanize(x.to_s)) if needs_headers
                row_data.push(x.is_a?(Symbol) ? instance.instance_eval(x.to_s) : x)
              end
            end

            data.push row_data
          end

        end
      end

      if headers
        if headers == true
          headers = false
        elsif !headers[0].is_a?(Array)
          headers = [headers]
        end
      end

      return options.merge(headers: headers, data: data, column_types: options[:column_types])
    end

    def self.get_options(options, klass)
      self.check_options_types(options)

      if options[:headers]
        if defined?(klass::SPREADSHEET_OPTIONS)
          header_style = deep_clone(SpreadsheetArchitect.default_options[:header_style]).merge(klass::SPREADSHEET_OPTIONS[:header_style] || {})
        else
          header_style = deep_clone(SpreadsheetArchitect.default_options[:header_style])
        end
        
        if options[:header_style]
          header_style.merge!(options[:header_style])
        elsif options[:header_style] == false
          header_style = false
        end
      else
        header_style = false
      end

      if options[:row_style] == false
        row_style = false
      else
        if defined?(klass::SPREADSHEET_OPTIONS)
          row_style = deep_clone(SpreadsheetArchitect.default_options[:row_style]).merge(klass::SPREADSHEET_OPTIONS[:row_style] || {})
        else
          row_style = deep_clone(SpreadsheetArchitect.default_options[:row_style])
        end

        if options[:row_style]
          row_style.merge!(options[:row_style])
        end
      end

      if defined?(klass::SPREADSHEET_OPTIONS)
        sheet_name = options[:sheet_name] || klass::SPREADSHEET_OPTIONS[:sheet_name] || SpreadsheetArchitect.default_options[:sheet_name]
      else
        sheet_name = options[:sheet_name] || SpreadsheetArchitect.default_options[:sheet_name]
      end

      sheet_name ||= (klass.name == 'SpreadsheetArchitect' ? 'Sheet1' : klass.name)

      return options.merge(header_style: header_style, row_style: row_style, sheet_name: sheet_name)
    end

    def self.convert_styles_to_ods(styles={})
      styles = {} unless styles.is_a?(Hash)
      styles = self.stringify_keys(styles)

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

    def self.deep_clone(x)
      Marshal.load(Marshal.dump(x))
    end

    def self.check_type(options, option_name, type)
      val = options[option_name]

      unless val.nil?
        invalid = false

        if type.is_a?(Array)
          invalid = type.all?{|t| !val.is_a?(t) }
        elsif !val.is_a?(type)
          invalid = true
        end

        if invalid
          raise SpreadsheetArchitect::Exceptions::IncorrectTypeError.new(option_name)
        end
      end
    end

    def self.check_options_types(options)
      self.check_type(options, :spreadsheet_columns, Array)
      self.check_type(options, :data, Array)
      self.check_type(options, :instances, Array)
      self.check_type(options, :headers, [TrueClass, FalseClass, Array])
      self.check_type(options, :header_style, Hash)
      self.check_type(options, :row_style, Hash)
      self.check_type(options, :column_styles, Array)
      self.check_type(options, :range_styles, Array)
      self.check_type(options, :merges, Array)
      self.check_type(options, :borders, Array)
      self.check_type(options, :column_widths, Array)
    end

    # only converts the first 2 levels
    def self.stringify_keys(hash={})
      new_hash = {}
      hash.each do |k,v|
        if v.is_a?(Hash)
          v.each do |k2,v2|
            new_hash[k2.to_s] = v2
          end
        else
          new_hash[k.to_s] = v
        end
      end
      return new_hash
    end

    def self.is_ar_model?(klass)
      defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
    end
  end
end
