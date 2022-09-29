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
    @path = TMP_PATH.join("models/#{klass}")
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

        save_file("models/#{klass}/instances.#{format}", data)
      end

      test "Empty :instances #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        data = which.send(method, instances: [])

        save_file("models/#{klass}/empty.#{format}", data)
      end

      test ":data #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        data = which.send(method, data: @data)

        save_file("models/#{klass}/data.#{format}", data)
      end

      test "Empty :data #{klass} #{format}" do
        set_path(klass)

        method = "to_#{format}"
        which = klass.respond_to?(method) ? klass : SpreadsheetArchitect

        data = which.send(method, data: [])

        save_file("models/#{klass}/empty.#{format}", data)
      end

      if klass.is_a?(ActiveRecord::Base)
        test "ActiveRecord::Relation" do
          method = "to_#{format}"

          data = klass.all.send(method)

          save_file("models/#{klass}/active_record_relation.#{format}", data)
        end

        test "Empty ActiveRecord::Relation" do
          method = "to_#{format}"

          data = klass.limit(0).send(method)

          save_file("models/#{klass}/empty_active_record_relation.#{format}", data)
        end
      end

    end

  end

end
