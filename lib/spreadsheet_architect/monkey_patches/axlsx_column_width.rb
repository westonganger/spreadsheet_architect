if RUBY_VERSION.to_f >= 2

  module SpreadsheetArchitect
    module AxlsxColumnWidthPatch
      def initialize(*args)
        @width = nil
        super
      end

      def width=(v)
        if v.nil?
          @custom_width = false
          @width = nil
        elsif @width.nil? || @width < v+5
          @custom_width = true
          @best_fit = true
          @width = v + 5
        else
          @custom_width = true
          @width = v
        end
      end
    end
  end

  module Axlsx
    class Col
      prepend ::SpreadsheetArchitect::AxlsxColumnWidthPatch
    end
  end

else

  module Axlsx
    class Col
      original_initialize = instance_method(:initialize)
      define_method :initialize do |*args|
        @width = nil

        original_initialize.bind(self).(*args)
      end

      def width=(v)
        if v.nil?
          @custom_width = false
          @width = nil
        elsif @width.nil? || @width < v+5
          @custom_width = true
          @best_fit = true
          @width = v + 5
        end
      end
    end
  end

end
