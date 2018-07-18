require "test_helper"

class AllModelsTest < ActiveSupport::TestCase

  def setup
    @data = [
      ['row1', 'test1'],
      [123.456, nil, '123'],
      [123, Date.today, Time.now]
    ]
  end

  def teardown
  end

  def set_path(klass)
    @path = VERSIONED_BASE_PATH.join("models/#{klass}")
    FileUtils.mkdir_p(@path)
  end

  models = []

  models_folder = File.expand_path('../../dummy_app/app/models', __FILE__)
  Dir[File.join(models_folder, '*.rb')].each do |filename|
    klass = filename.split('/').last.gsub('.rb', '').titleize.gsub(' ','').constantize
    unless [ApplicationRecord].include?(klass)
      models.push(klass)
    end
  end

  models.each do |klass|
    instances = 5.times.map{|i| 
      x = (klass == SpreadsheetArchitect ? Post : klass).new
      x.name = i
      x.content = i+2
      x.created_at = Time.now
      x
    }

    ['csv', 'ods', 'xlsx'].each do |format|

      test ":instances #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        data = which.send(method, instances: instances)

        File.open(File.join(@path, "instances.#{format}"), 'w+b') do |f|
          f.write data
        end
      end

      test "Empty :instances #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        File.open(File.join(@path, "empty.#{format}"),'w+b') do |f|
          f.write which.send(method, instances: [])
        end
      end

      test ":data #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        data = which.send(method, data: @data)

        File.open(File.join(@path, "data.#{format}"), 'w+b') do |f|
          f.write data
        end
      end

      test "Empty :data #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        File.open(File.join(@path, "empty.#{format}"),'w+b') do |f|
          f.write which.send(method, data: [])
        end
      end

      if klass.is_a?(ActiveRecord::Base)
        test "ActiveRecord::Relation" do
          method = "to_#{format}"

          File.open(File.join(@path, "active_record_relation.#{format}"),'w+b') do |f|
            f.write klass.all.send(method)
          end
        end

        test "Empty ActiveRecord::Relation" do
          method = "to_#{format}"

          File.open(File.join(@path, "empty_active_record_relation.#{format}"),'w+b') do |f|
            f.write klass.limit(0).send(method)
          end
        end
      end

    end

  end

end
