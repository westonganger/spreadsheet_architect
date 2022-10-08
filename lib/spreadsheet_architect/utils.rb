module SpreadsheetArchitect
  module Utils
    def self.get_cell_data(options, klass)
      if options[:data] && options[:instances]
        raise SpreadsheetArchitect::Exceptions::MultipleDataSourcesError
      elsif options[:data]
        data = options[:data]
      end

      if options[:headers] == true
        headers = []

        if options[:data]
          headers = false
        else
          needs_headers = true
        end
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

      if !options[:skip_defaults]
        if defined?(klass::SPREADSHEET_OPTIONS)
          if klass::SPREADSHEET_OPTIONS.is_a?(Hash)
            defaults = SpreadsheetArchitect.default_options.merge(klass::SPREADSHEET_OPTIONS)
          else
            raise SpreadsheetArchitect::Exceptions::OptionTypeError.new("#{klass}::SPREADSHEET_OPTIONS constant")
          end
        else
          defaults = SpreadsheetArchitect.default_options
        end

        options = defaults.merge(options)
      end

      if !options[:headers]
        options.delete(:header_style)
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

      if options[:column_types] && !(options[:column_types].compact.collect(&:to_sym) - SpreadsheetArchitect::XLSX_COLUMN_TYPES).empty?
        raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Invalid column type. Valid XLSX values are #{SpreadsheetArchitect::XLSX_COLUMN_TYPES}")
      end

      if options[:freeze]
        options[:freeze] = SpreadsheetArchitect::Utils.symbolize_keys(options[:freeze])

        if options[:freeze_headers]
          raise SpreadsheetArchitect::Exceptions::ArgumentError.new('Cannot use both :freeze and :freeze_headers options at the same time')
        end
      end

      if options[:escape_formulas].nil?
        options[:escape_formulas] = true
      end

      if options[:use_zero_based_row_index].nil?
        options[:use_zero_based_row_index] = false
      end

      return options
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
      options = self.symbolize_keys(options, shallow: true)

      bad_keys = options.keys - ALLOWED_OPTIONS.keys

      if bad_keys.any?
        raise SpreadsheetArchitect::Exceptions::ArgumentError.new("Invalid options provided: #{bad_keys}")
      end

      ALLOWED_OPTIONS.each do |key, allowed_types|
        check_option_type(options, key, allowed_types)
      end
    end

    def self.stringify_keys(hash)
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

    def self.symbolize_keys(hash, shallow: false)
      new_hash = {}

      hash.each do |k,v|
        if v.is_a?(Hash)
          new_hash[k.to_sym] = shallow ? v : self.symbolize_keys(v)
        else
          new_hash[k.to_sym] = v
        end
      end

      return new_hash
    end

    def self.hash_array_symbolize_keys(array)
      new_array = []

      array.each_with_index do |x,i|
        new_array[i] = x.is_a?(Hash) ? self.symbolize_keys(x) : x
      end

      return new_array
    end

    ALLOWED_OPTIONS = {
      borders: Array,
      column_styles: Array,
      conditional_row_styles: Array,
      column_widths: Array,
      column_types: Array,
      data: Array,
      freeze_headers: [TrueClass, FalseClass],
      freeze: Hash,
      headers: [TrueClass, FalseClass, Array],
      header_style: Hash,
      instances: Array,
      merges: Array,
      range_styles: Array,
      skip_defaults: [TrueClass, FalseClass],
      row_style: Hash,
      sheet_name: String,
      spreadsheet_columns: [Proc, Symbol, String],
      escape_formulas: [TrueClass, FalseClass, Array],
      use_zero_based_row_index: [TrueClass, FalseClass],
    }.freeze

  end
end
