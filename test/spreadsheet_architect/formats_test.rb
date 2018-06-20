require 'test_helper'

class FormatsTest < ActiveSupport::TestCase

  [:xlsx, :ods, :csv].each do |format|
    test "Registers :#{format} mime type" do
      assert Mime::Type.lookup_by_extension(format)
    end

    test "Registers :#{format} template handler" do
      assert ActionController::Renderers::RENDERERS.include?(format.to_sym)
    end
  end

end
