source 'https://rubygems.org'

gemspec

### Must remain in the Gemfile, not gemspec, otherwise Rails integration tests fail
gem 'responders'

def get_env(name)
  (ENV[name] && !ENV[name].empty?) ? ENV[name] : nil
end

gem "rails", get_env("RAILS_VERSION")

db_gem = get_env("DB_GEM") || "sqlite3"
gem db_gem, get_env("DB_GEM_VERSION")

gem "minitest-spec-rails"
