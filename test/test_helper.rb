ENV["RAILS_ENV"] = "test"

require File.expand_path("../dummy_app/config/environment.rb",  __FILE__)

migration_path = Rails.root.join('db/migrate')
if ActiveRecord.gem_version >= ::Gem::Version.new("5.2.0")
  ActiveRecord::MigrationContext.new(migration_path).migrate
else
  ActiveRecord::Migrator.migrate(migration_path)
end

require "rails/test_help"

Rails.backtrace_cleaner.remove_silencers!

class ActiveSupport::TestCase
  # Setup all fixtures in test/fixtures/*.yml for all tests in alphabetical order.
  fixtures :all
end

require 'minitest/reporters'
Minitest::Reporters.use!(
  Minitest::Reporters::DefaultReporter.new,
  ENV,
  Minitest.backtrace_filter
)

require 'custom_assertions'

post_count = Post.count
if post_count < 5
  (5 - post_count).times do |i|
    Post.create!(name: "foo #{i}", content: "bar #{i}", age: i)
  end
end

axlsx_spec = Gem.loaded_specs['axlsx']
if !axlsx_spec.source.is_a?(Bundler::Source::Rubygems)
  axlsx_version = 'axlsx-master'
else
  axlsx_version = axlsx_spec.version.to_s
end

VERSIONED_BASE_PATH = Rails.root.join("tmp/#{axlsx_version}")

### Cleanup old test spreadsheets
FileUtils.remove_dir(VERSIONED_BASE_PATH, true)
FileUtils.mkdir_p(VERSIONED_BASE_PATH)
