lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'spreadsheet_architect/version.rb'

Gem::Specification.new do |s|
  s.name        = 'spreadsheet_architect'
  s.version     =  SpreadsheetArchitect::VERSION
  s.author	= "Weston Ganger"
  s.email       = 'westonganger@gmail.com'
  s.homepage 	= 'https://github.com/westonganger/spreadsheet_architect'
  
  s.summary     = "Spreadsheet Generator for ActiveRecord Models and Ruby Classes/Modules"
  s.description = "Spreadsheet Architect lets you turn any activerecord relation or ruby object collection into a XLSX, ODS, or CSV spreadsheets. Generates columns from model activerecord column_names or from an array of ruby methods."
  s.files = Dir.glob("{lib/**/*}") + %w{ LICENSE README.md Rakefile CHANGELOG.md }
  s.test_files  = Dir.glob("{test/**/*}")

  s.add_runtime_dependency 'axlsx', '>= 2.0'
  s.add_runtime_dependency 'axlsx_styler', '>= 0.1.6'
  s.add_runtime_dependency 'rodf', '>= 1.0.0'
  
  s.add_development_dependency 'rake'
  s.add_development_dependency 'minitest'
  s.add_development_dependency 'bundler'
  s.add_development_dependency 'sqlite3'
  s.add_development_dependency 'activerecord'

  s.required_ruby_version = '>= 1.9.3'
  s.require_path = 'lib'
end
