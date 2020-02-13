if Axlsx::VERSION.to_f < 3 && (Axlsx::VERSION.to_s.split('.')[2].to_i < 2)

  Axlsx::Cell.class_eval do
    private

    def string_width(string, font_size)
      font_scale = font_size / 10.0
      (string.to_s.size + 3) * font_scale
    end
  end

  Axlsx::RichTextRun.class_eval do
    private

    def string_width(string, font_size)
      font_scale = font_size / 10.0
      string.to_s.size * font_scale
    end
  end

end
