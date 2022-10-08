module SpreadsheetArchitect
  module Utils
    module ODS

      def self.get_cell_type(value, type=nil)
        if type && !type.empty?
          case type
          when :hyperlink
            return :string
          when :date, :time
            return :string
          end

          return type unless (type.respond_to?(:empty?) ? type.empty? : type.nil?)
        end
        
        if value.is_a?(Numeric)
          type = :float
        elsif value.respond_to?(:strftime)
          type = :string
        else
          type = :string
        end

        return type
      end

      def self.convert_styles(styles={})
        styles = {} unless styles.is_a?(Hash)
        styles = SpreadsheetArchitect::Utils.stringify_keys(styles)

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

    end
  end
end
