module SpreadsheetArchitect
  module Utils
    def self.str_humanize(str, capitalize = true)
      str = str.sub(/\A_+/, '').gsub(/[_\.]/,' ').sub(' rescue nil','')
      if capitalize
        str = str.gsub(/(\A|\ )\w/){|x| x.upcase}
      end
      return str
    end

    def self.get_cell_data(options={}, klass)
      self.check_options_types

      if klass.name == 'SpreadsheetArchitect'
        if !options[:data]
          raise SpreadsheetArchitect::Exceptions::NoDataError
        end

        if options[:headers] && options[:headers].is_a?(Array) && !options[:headers].empty?
          headers = options[:headers]
        else
          headers = false
        end
        
        data = options[:data]
      else
        has_custom_columns = options[:spreadsheet_columns] || klass.instance_methods.include?(:spreadsheet_columns)

        if !options[:instances] && defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
          options[:instances] = klass.where(nil).to_a # triggers the relation call, not sure how this works but it does
        end

        if !options[:instances]
          raise SpreadsheetArchitect::Exceptions::NoInstancesError
        end

        if options[:headers].nil?
          headers = klass::SPREADSHEET_OPTIONS[:headers] if defined?(klass::SPREADSHEET_OPTIONS)
          headers ||= SpreadsheetArchitect.default_options[:headers]
        elsif options[:headers].is_a?(Array)
          headers = options[:headers]
        end

        if headers == false || headers.is_a?(Array)
          needs_headers = false
        else
          headers = true
          needs_headers = true
        end

        if has_custom_columns 
          headers = [] if needs_headers
          columns = []
          array = options[:spreadsheet_columns] || (options[:instances].empty? ? [] : options[:instances].first.spreadsheet_columns)
          array.each_with_index do |x,i|
            if x.is_a?(Array)
              headers.push(x[0].to_s) if needs_headers
              columns.push x[1]
            else
              headers.push(str_humanize(x.to_s)) if needs_headers
              columns.push x
            end
          end
        elsif !has_custom_columns && defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
          ignored_columns = ["id","created_at","updated_at","deleted_at"] 
          the_column_names = (klass.column_names - ignored_columns)
          headers = the_column_names.map{|x| str_humanize(x)} if needs_headers
          columns = the_column_names.map{|x| x.to_sym}
        else
          raise SpreadsheetArchitect::Exceptions::SpreadsheetColumnsNotDefinedError, klass
        end

        data = []
        options[:instances].each do |instance|
          if has_custom_columns && !options[:spreadsheet_columns]
            row_data = []
            instance.spreadsheet_columns.each do |x|
              if x.is_a?(Array)
                row_data.push(x[1].is_a?(Symbol) ? instance.instance_eval(x[1].to_s) : x[1])
              else
                row_data.push(x.is_a?(Symbol) ? instance.instance_eval(x.to_s) : x)
              end
            end
            data.push row_data
          else
            data.push columns.map{|col| col.is_a?(Symbol) ? instance.instance_eval(col.to_s) : col}
          end
        end
      end

      if headers && !headers[0].is_a?(Array)
        headers = [headers]
      end

      return options.merge(headers: headers, data: data, column_types: options[:column_types])
    end

    def self.get_options(options={}, klass)
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
      unless options[option_name].nil?
        valid = false

        if type.is_a?(Array)
          valid = type.any?{|t| options[option_name].is_a?(t)}
        elsif options[option_name].is_a?(type)
          valid = true
        end

        if valid
          raise SpreadsheetArchitect::Exceptions::IncorrectTypeError option_name
        end
      end
    end

    def self.check_options_types(options={})
      self.check_type(options, :spreadsheet_columns, Array)
      self.check_type(options, :instances, Array)
      self.check_type(options, :headers, [TrueClass, FalseClass, Array])
      self.check_type(options, :sheet_name, String)
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
  end
end
