source 'https://rubygems.org'

gemspec

### Must remain in the Gemfile, not gemspec, otherwise Rails integration tests fail
gem 'responders'

def get_env(name)
  (ENV[name] && !ENV[name].empty?) ? ENV[name] : nil
end

gem "rails", get_env("RAILS_VERSION")
gem "sqlite3"

gem "minitest-spec-rails", git: "https://github.com/metaskills/minitest-spec-rails.git"
