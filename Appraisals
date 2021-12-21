appraise "caxlsx" do
  gem "caxlsx" # Ruby 2.3+
end

appraise "caxlsx_3.0.0" do
  gem "caxlsx", "3.0.0" # Legacy Axlsx Support
end

appraise "caxlsx_2.0.2" do
  gem "caxlsx", "2.0.2" # Legacy Axlsx Support
end

["sqlite3"].each do |db_gem|
  appraise "rails_7.0.#{db_gem}" do
    gem "rails", "~> 7.0.0"
    gem 'responders'
    gem db_gem
  end

  appraise "rails_6.1.#{db_gem}" do
    gem "rails", "~> 6.1.1"
    gem 'responders'
    gem db_gem
  end

  appraise "rails_6.0.#{db_gem}" do
    gem "rails", "~> 6.0.3"
    gem 'responders'
    gem db_gem
  end

  appraise "rails_5.2.#{db_gem}" do
    gem "rails", "~> 5.2.4"
    gem 'responders'
    gem db_gem
  end

  appraise "rails_5.1.#{db_gem}" do
    gem "rails", "~> 5.1.7"
    gem 'responders'
    gem db_gem
  end

  appraise "rails_5.0.#{db_gem}" do
    gem "rails", "~> 5.0.7"
    gem 'responders'

    if db_gem == 'sqlite3'
      gem "sqlite3", "~> 1.3.13"
    else
      gem db_gem
    end
  end
end
