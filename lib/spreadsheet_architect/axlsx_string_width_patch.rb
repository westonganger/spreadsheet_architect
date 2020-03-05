if Axlsx::VERSION.to_f < 3.1

  Axlsx::Cell.class_eval do
    private

    def string_width(string, font_size)
      font_scale = font_size / 10.0
      (string.to_s.size + 3) * font_scale
    end
  end

  if defined?(Axlsx::RichTextRun)
    Axlsx::RichTextRun.class_eval do
      private

      def string_width(string, font_size)
        font_scale = font_size / 10.0
        string.to_s.size * font_scale
      end
    end
  end

end
