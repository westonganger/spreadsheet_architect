if defined? Axlsx
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
          @custom_width = @best_fit = v != nil
          @width = v + 5
        end
      end

    end
  end
end
