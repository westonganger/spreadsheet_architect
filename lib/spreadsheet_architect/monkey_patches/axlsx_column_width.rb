module Axlsx
  class Cell
    def string_width(string, font_size)
      font_scale = font_size / 10.0
      (string.to_s.size + 3) * font_scale
    end
  end
end
