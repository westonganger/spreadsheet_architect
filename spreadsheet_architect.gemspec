lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'spreadsheet_architect/version.rb'

Gem::Specification.new do |s|
  s.name        = 'spreadsheet_architect'
  s.version     =  SpreadsheetArchitect::VERSION
  s.author	= "Weston Ganger"
  s.email       = 'weston@westonganger.com'
  s.homepage 	= 'https://github.com/westonganger/spreadsheet_architect'
  
  s.summary     = "Spreadsheet Architect is a library that allows you to create XLSX, ODS, or CSV spreadsheets easily from ActiveRecord relations, Plain Ruby classes, or predefined data."
  s.description = s.summary 

  s.files = Dir.glob("{lib/**/*}") + %w{ LICENSE README.md Rakefile CHANGELOG.md }
  s.test_files  = Dir.glob("{test/**/*}")
  s.require_path = 'lib'

  s.required_ruby_version = '>= 2.3.0'

  s.add_runtime_dependency 'caxlsx', ['>= 2.0.2', '<4']
  s.add_runtime_dependency 'axlsx_styler', ['>= 1.0.0', '<2']
  s.add_runtime_dependency 'rodf', ['>= 1.0.0', '<2']
  
  s.add_development_dependency 'rake'
  s.add_development_dependency 'bundler'
  s.add_development_dependency 'minitest'
  s.add_development_dependency 'minitest-reporters'
  s.add_development_dependency 'appraisal'
  s.add_development_dependency 'sqlite3'
  s.add_development_dependency 'rails'
end
