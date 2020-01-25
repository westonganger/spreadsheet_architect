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
