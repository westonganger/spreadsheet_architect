require File.expand_path('../boot', __FILE__)

require 'rails/all'

Bundler.require

require "spreadsheet_architect"

module Dummy
  class Application < Rails::Application
    # Settings in config/environments/* take precedence over those specified here.
    # Application configuration should go into files in config/initializers
    # -- all .rb files in that directory are automatically loaded.

    # Custom directories with classes and modules you want to be autoloadable.
    # config.autoload_paths += %W(#{config.root}/extras)

    # Only load the plugins named here, in the order given (default is alphabetical).
    # :all can be used as a placeholder for all plugins not explicitly named.
    # config.plugins = [ :exception_notification, :ssl_requirement, :all ]

    # Activate observers that should always be running.
    # config.active_record.observers = :cacher, :garbage_collector, :forum_observer

    # Set Time.zone default to the specified zone and make Active Record auto-convert to this zone.
    # Run "rake -D time" for a list of tasks for finding time zone names. Default is UTC.
    # config.time_zone = 'Central Time (US & Canada)'

    # The default locale is :en and all translations from config/locales/*.rb,yml are auto loaded.
    # config.i18n.load_path += Dir[Rails.root.join('my', 'locales', '*.{rb,yml}').to_s]
    # config.i18n.default_locale = :de

    # Configure the default encoding used in templates for Ruby 1.9.
    config.encoding = "utf-8"

    # Configure sensitive parameters which will be filtered from the log file.
    config.filter_parameters += [:password]

    config.generators.test_framework = false
    config.generators.helper = false
    config.generators.stylesheets = false
    config.generators.javascripts = false

    config.after_initialize do
      ActiveRecord::Migration.migrate(Rails.root.join("db/migrate/*").to_s)
    end

    if ActiveRecord.respond_to?(:gem_version)
      gem_version = ActiveRecord.gem_version

      if gem_version >= Gem::Version.new("7.0.0")
        config.active_record.legacy_connection_handling = false
      elsif gem_version.to_s.start_with?("5.2.")
        config.active_record.sqlite3.represent_boolean_as_integer = true
      end
    end
  end
end
