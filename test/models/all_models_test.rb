require "test_helper"

class AllModelsTest < ActiveSupport::TestCase

  def setup
  end

  def teardown
  end

  def set_path(klass)
    @path = Rails.root.join("tmp/configurations/#{klass}")
    FileUtils.mkdir_p(@path)
  end

  [ActiveModelObject, PlainRubyObject, LegacyPlainRubyObject, Post, CustomPost].each do |klass|
    instances = 5.times.map{|i| 
      x = klass.new
      x.name = i
      x.content = i+2
      x.created_at = Time.now
      x
    }

    ['csv', 'ods', 'xlsx'].each do |format|

      test "#{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = (!klass.is_a?(ActiveRecord::Base) && klass.respond_to?(method)) ? klass : SpreadsheetArchitect

        data = which.send(method, instances: instances)

        File.open(File.join(@path, '#{format}.#{format}'), 'w+b') do |f|
          f.write data
        end
      end

    end

  end

end
