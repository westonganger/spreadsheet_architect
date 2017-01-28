ENV["RAILS_ENV"] = "test"

require File.expand_path("../../config/environment", __FILE__)

ActiveRecord::Migrator.migrate("db/migrate")

require "rails/test_help"
require "minitest/rails"
require 'minitest/pride'
require 'minitest/reporters'
require 'database_cleaner'

Minitest::Reporters.use!
DatabaseCleaner.strategy = :transaction

class ActiveSupport::TestCase
  # Setup all fixtures in test/fixtures/*.yml for all tests in alphabetical order.
  fixtures :all

  # Add more helper methods to be used by all tests here...
end
