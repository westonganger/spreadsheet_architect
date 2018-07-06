module SpreadsheetArchitect
  module Utils
    module XLSX

      def self.get_type(value, type=nil)
        return type unless (type.respond_to?(:empty?) ? type.empty? : type.nil?)
        
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
        styles = {} unless styles.is_a?(Hash)
        styles = self.symbolize_keys(styles)

        if styles[:fg_color].nil? && styles[:color] && styles[:color].respond_to?(:sub) && !styles[:color].empty?
          styles[:fg_color] = styles.delete(:color).sub('#','')
        end

        if styles[:bg_color].nil? && styles[:background_color] && styles[:background_color].respond_to?(:sub) && !styles[:background_color].empty?
          styles[:bg_color] = styles.delete(:background_color).sub('#','')
        end
        
        if styles[:alignment].nil? && styles[:align]
          if styles[:align].is_a?(Hash)
            styles[:alignment] = {
              horizontal: styles[:align][:horizontal], 
              vertical: styles[:align][:vertical], 
              wrap_text: styles[:align][:wrap_text]
            }
          else
            styles[:alignment] = {horizontal: (styles.delete(:align) || nil) }
          end

          styles.delete(:align)
        end

        if styles[:b].nil?
          styles[:b] = styles.delete(:bold) || nil
        end

        if styles[:sz].nil?
          styles[:sz] = styles.delete(:font_size) || nil
        end

        if styles[:i].nil?
          styles[:i] = styles.delete(:italic) || nil
        end

        if styles[:u].nil?
          styles[:u] = styles.delete(:underline) || nil
        end

        ### If `:u` is false instead of nil, it may be incorrectly rendered as true in Excel
        if styles[:u] == false
          styles[:u] = nil
        end
        
        styles.delete_if{|k,v| v.nil?}
      end

      def self.range_hash_to_str(hash, num_columns, num_rows)
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
        when :all
          start_col = 'A'
          end_col = COL_NAMES[num_columns-1]
        else
          raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError.new(:columns, hash)
        end

        case hash[:rows]
        when Integer
          start_row = end_row = hash[:rows]
        when Range
          start_row = hash[:rows].first
          end_row = hash[:rows].last
        when :all
          start_row = 1
          end_row = num_rows
        else
          raise SpreadsheetArchitect::Exceptions::InvalidRangeStylesOptionError.new(:rows, hash)
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
              raise SpreadsheetArchitect::Exceptions::InvalidRangeError.new(:columns, range)
            end
            
            unless start_row.to_i <= num_rows && end_row.to_i <= num_rows
              raise SpreadsheetArchitect::Exceptions::InvalidRangeError.new(:rows, range)
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

      private

      def self.symbolize_keys(hash={})
        new_hash = {}
        hash.each do |k, v|
          if v.is_a?(Hash)
            new_hash[k.to_sym] = self.symbolize_keys(v)
          else
            new_hash[k.to_sym] = v
          end
        end
        new_hash
      end

      ### Limit of 16384 columns as per Excel limitations
      COL_NAMES = Array('A'..'XFD').freeze

    end
  end
end
