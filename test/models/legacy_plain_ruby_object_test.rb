require "test_helper"

class LegacyPlainRubyObjectTest < ActiveSupport::TestCase

  def setup
    @klass = LegacyPlainRubyObject

    @instances = 5.times.map{|i| 
      x = @klass.new
      x.name = i
      x.content = i+2
      x.created_at = Time.now
      x
    }
  end

  def teardown
  end

  test "Is not an AR object" do
    assert_not @instances.first.is_a?(ActiveRecord::Base)
  end

  ['csv', 'ods', 'xlsx'].each do |format|

    test "Legacy PORO #{format}" do
      @klass.to_csv(instances: @instances)
    end

  end

end
