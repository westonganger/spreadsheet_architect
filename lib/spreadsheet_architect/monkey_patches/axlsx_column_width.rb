if defined? Axlsx
  Axlsx::Col.class_eval do
    @width = nil

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
