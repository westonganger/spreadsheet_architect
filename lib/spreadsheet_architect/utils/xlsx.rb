module SpreadsheetArchitect
  module Utils
    module XLSX

      def self.get_type(value, type=nil)
        if type && !type.empty?
          case type
          when :hyperlink
            return :string
          end

          return type unless (type.respond_to?(:empty?) ? type.empty? : type.nil?)
        end
        
        if value.is_a?(Numeric)
          if value.is_a?(Float) || value.is_a?(BigDecimal)
            type = :float
          else
            type = :integer
          end
        elsif value.respond_to?(:strftime)
          if value.is_a?(DateTime) || value.is_a?(Time)
            type = :time
          else
            type = :date
          end
        else
          type = :string
        end

        return type
      end

      def self.convert_styles_to_axlsx(styles={})
        if [nil, false, true].include?(styles)
          return {}
        end

        styles = SpreadsheetArchitect::Utils.symbolize_keys(styles)

        ### BOOLEAN VALUES
        if styles[:b].nil? && styles.has_key?(:bold)
          styles[:b] = !!styles.delete(:bold)
        end

        if styles[:i].nil? && styles.has_key?(:italic)
          styles[:i] = !!styles.delete(:italic)
        end

        if styles[:u].nil? && styles.has_key?(:underline)
          styles[:u] = !!styles.delete(:underline)
        end

        ### OTHER VALUES
        if styles[:sz].nil? && !styles[:font_size].nil?
          styles[:sz] = styles.delete(:font_size)
        end

        if styles[:fg_color].nil? && styles[:color] && styles[:color].respond_to?(:sub) && !styles[:color].empty?
          styles[:fg_color] = styles.delete(:color).sub('#','')
        end

        if styles[:bg_color].nil? && styles[:background_color] && styles[:background_color].respond_to?(:sub) && !styles[:background_color].empty?
          styles[:bg_color] = styles.delete(:background_color).sub('#','')
        end
        
        if styles[:alignment].nil? && [:align, :valign, :wrap_text].any?{|k| styles.has_key?(k) }
          if styles[:align].is_a?(Hash)
            styles[:alignment] = {
              horizontal: styles[:align][:horizontal], 
              vertical: styles[:align][:vertical], 
              wrap_text: styles[:align][:wrap_text]
            }
          else
            styles[:alignment] = {horizontal: styles.delete(:align), vertical: styles.delete(:valign), wrap_text: styles.delete(:wrap_text) }.compact
          end
        end

        ### COMMENT SEEMS WRONG, TO BE RECONFIRMED
        ### If `:u` is false instead of nil, it may be incorrectly rendered as true in Excel
        #if styles[:u] == false
        #  styles[:u] = nil
        #end
        
        ### ENSURE CLEANUP OF ALL ALIAS KEYS
        [:bold, :font_size, :italic, :underline, :align, :valign, :wrap_text, :color, :background_color].each do |k|
          styles.delete(k)
        end
        
        return styles
      end

      def self.range_hash_to_str(hash, num_columns, num_rows, use_zero_based_row_index: false)
        if !hash.has_key?(:rows) && !hash.has_key?(:columns)
          raise SpreadsheetArchitect::Exceptions::InvalidRangeError.new(:missing_range_keys, hash)
        end

        case hash[:columns]
        when Integer
          start_col = end_col = COL_NAMES[hash[:columns]]
        when String
          start_col = hash[:columns].first
          end_col = hash[:columns].last
        when Range
          start_col = hash[:columns].first
          unless start_col.is_a?(String)
            start_col = COL_NAMES[start_col]
          end

          end_col = hash[:columns].last
          unless end_col.is_a?(String)
            end_col = COL_NAMES[end_col]
          end
        when :all, nil
          start_col = 'A'
          end_col = COL_NAMES[num_columns-1]
        else
          raise SpreadsheetArchitect::Exceptions::InvalidRangeOptionError.new(:columns, hash)
        end

        case hash[:rows]
        when Integer
          start_row = end_row = hash[:rows]

          if use_zero_based_row_index
            start_row += 1
            end_row += 1
          end
        when Range
          start_row = hash[:rows].first
          end_row = hash[:rows].last

          if use_zero_based_row_index
            start_row += 1
            end_row += 1
          end
        when :all, nil
          start_row = 1
          end_row = num_rows
        else
          raise SpreadsheetArchitect::Exceptions::InvalidRangeOptionError.new(:rows, hash)
        end

        range_str = "#{start_col}#{start_row}:#{end_col}#{end_row}"

        unless hash[:columns] == :all && hash[:rows] == :all
          verify_range(range_str, num_rows)
        end

        return range_str
      end

      def self.verify_range(range, num_rows)
        if range.is_a?(String)
          if range.include?(':')
            front, back = range.split(':')
            start_col, start_row = front.scan(/\d+|\D+/)
            end_col, end_row = back.scan(/\d+|\D+/)

            unless COL_NAMES.include?(start_col) && COL_NAMES.include?(end_col)
              raise SpreadsheetArchitect::Exceptions::InvalidRangeValue.new(:columns, range)
            end
            
            unless start_row.to_i <= num_rows && end_row.to_i <= num_rows
              raise SpreadsheetArchitect::Exceptions::InvalidRangeValue.new(:rows, range)
            end
          else
            raise SpreadsheetArchitect::Exceptions::InvalidRangeError.new(:format, range)
          end
        else
          raise SpreadsheetArchitect::Exceptions::InvalidRangeError.new(:type, range)
        end
      end

      def self.verify_column(col, num_columns)
        unless (col.is_a?(String) && COL_NAMES.include?(col)) || (col.is_a?(Integer) && col >= 0 && col < num_columns)
          raise SpreadsheetArchitect::Exceptions::InvalidColumnError.new(col)
        end
      end

      def self.conditional_styles_for_row(conditional_row_styles, row_index, row_data)
        merged_conditional_styles = {}

        conditional_row_styles.each do |x|
          if x[:if] && x[:unless]
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new('Cannot pass both :if and :unless within the same :conditonal_row_styles entry')
          elsif !x[:if] && !x[:unless]
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new('Must pass either :if or :unless within the each :conditonal_row_styles entry')
          elsif !x[:styles]
            raise SpreadsheetArchitect::Exceptions::ArgumentError.new('Must pass the :styles option within a :conditonal_row_styles entry')
          end

          conditions_met = (x[:if] || x[:unless]).call(row_data, row_index)

          if x[:unless]
            conditions_met = !conditions_met
          end

          if conditions_met
            merged_conditional_styles.merge!(x[:styles])
          end
        end

        return merged_conditional_styles
      end

      private

      ### Limit of 16384 columns as per Excel limitations
      COL_NAMES = Array('A'..'XFD').freeze

    end
  end
end
