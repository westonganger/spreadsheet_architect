module SpreadsheetArchitect
  module Utils
    def self.str_humanize(str, capitalize = true)
      str = str.sub(/\A_+/, '').gsub(/[_\.]/,' ').sub(' rescue nil','')
      if capitalize
        str = str.gsub(/(\A|\ )\w/){|x| x.upcase}
      end
      return str
    end

    def self.get_type(value, type=nil, last_run=false)
      return type if !type.blank?
      if value.is_a?(Numeric)
        if [:float, :decimal].include?(type)
          type = :float
        else
          type = :integer
        end
      elsif value.is_a?(Date)
        type = :date
      elsif value.is_a?(DateTime) || value.is_a?(Time)
        type = :time
      elsif !last_run && value.is_a?(Symbol)
        type = :symbol
      else
        type = :string
      end
      return type
    end

    def self.get_cell_data(options={}, klass)
      if klass.name == 'SpreadsheetArchitect'
        if !options[:data] || options[:data].empty?
          raise SpreadsheetArchitect::NoDataError
        end

        if options[:headers] && !options[:headers].empty?
          headers = options[:headers]
        else
          headers = false
        end
        
        data = options[:data]

        types = []
        data.first.each do |x|
          types.push self.get_type(x, nil)
        end
      else
        has_custom_columns = options[:spreadsheet_columns] || klass.instance_methods.include?(:spreadsheet_columns)

        if !options[:instances] && defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
          options[:instances] = klass.where(nil).to_a # triggers the relation call, not sure how this works but it does
        end

        if !options[:instances] || options[:instances].empty?
          raise SpreadsheetArchitect::NoInstancesError
        end

        if options[:headers].nil?
          headers = klass::SPREADSHEET_OPTIONS[:headers] if defined?(klass::SPREADSHEET_OPTIONS)
          headers ||= SpreadsheetArchitect.default_options[:headers]
        else
          headers = options[:headers]
        end

        unless headers == false || headers.is_a?(Array)
          headers = klass.spreadsheet_headers if klass.respond_to?(:spreadsheet_headers)
        end

        if has_custom_columns 
          headers = [] if headers.nil?
          columns = []
          types = []
          array = options[:spreadsheet_columns] || options[:instances].first.spreadsheet_columns
          array.each do |x|
            if x.is_a?(Array)
              headers.push(x[0].to_s) if headers.nil?
              columns.push x[1]
              types.push self.get_type(x[1], nil)
            else
              headers.push(str_humanize(x.to_s)) if headers.nil?
              columns.push x
              types.push self.get_type(x, nil)
            end
          end
        elsif !has_custom_columns && defined?(ActiveRecord) && klass.ancestors.include?(ActiveRecord::Base)
          ignored_columns = ["id","created_at","updated_at","deleted_at"] 
          the_column_names = (klass.column_names - ignored_columns)
          headers = the_column_names.map{|x| str_humanize(x)} if headers.nil?
          columns = the_column_names.map{|x| x.to_sym}
          types = klass.columns.keep_if{|x| !ignored_columns.include?(x.name)}.collect(&:type)
          types.map!{|type| self.get_type(nil, type)}
        else
          raise SpreadsheetArchitect::SpreadsheetColumnsNotDefined, klass
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
      
        # Fixes missing types from symbol methods
        if has_custom_columns || options[:spreadsheet_columns]
          data.first.each_with_index do |x,i|
            if types[i] == :symbol
              types[i] = self.get_type(x, nil, true)
            end
          end
        end
      end

      return options.merge(headers: headers, data: data, types: types)
    end

    def self.get_options(options={}, klass)
      if options[:headers]
        if defined?(klass::SPREADSHEET_OPTIONS)
          header_style = SpreadsheetArchitect.default_options[:header_style].merge(klass::SPREADSHEET_OPTIONS[:header_style] || {})
        else
          header_style = SpreadsheetArchitect.default_options[:header_style]
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
          row_style = SpreadsheetArchitect.default_options[:row_style].merge(klass::SPREADSHEET_OPTIONS[:row_style] || {})
        else
          row_style = SpreadsheetArchitect.default_options[:row_style]
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

    def self.convert_styles_to_axlsx(styles={})
      styles[:fg_color] = styles.delete(:color) || styles[:fg_color]
      styles[:bg_color] = styles.delete(:background_color) || styles[:bg_color]
      if styles[:align].present?
        styles[:alignment] = {horizontal: styles.delete(:align)}
      end
      styles[:b] = styles.delete(:bold) || styles[:b]
      styles[:sz] = styles.delete(:font_size) || styles[:sz]
      styles[:i] = styles.delete(:italic) || styles[:i]
      styles[:u] = styles.delete(:underline) || styles[:u]

      styles.delete_if{|k,v| v.nil?}
    end
  end
end
