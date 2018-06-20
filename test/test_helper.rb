ENV["RAILS_ENV"] = "test"

require File.expand_path("../dummy_app/config/environment.rb",  __FILE__)
require "rails/test_help"

Rails.backtrace_cleaner.remove_silencers!

migration_path = Rails.root.join('db/migrate')
if ActiveRecord.gem_version >= ::Gem::Version.new("5.2.0")
  ActiveRecord::MigrationContext.new(migration_path).migrate
else
  ActiveRecord::Migrator.migrate(migration_path)
end
