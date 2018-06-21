require "test_helper"

class LegacyPlainRubyObjectTest < ActiveSupport::TestCase

  def setup
  end

  def test_csv
    LegacyPlainRubyObject.to_csv(instances: [])
  end

  def test_ods
    LegacyPlainRubyObject.to_ods(instances: [])
  end

  def test_xlsx
    LegacyPlainRubyObject.to_xlsx(instances: [])
  end

  def teardown
  end
end
