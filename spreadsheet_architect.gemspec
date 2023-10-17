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
  s.license     = 'MIT'

  s.files = Dir.glob("{lib/**/*}") + %w{ LICENSE README.md Rakefile CHANGELOG.md }
  s.require_path = 'lib'

  s.add_runtime_dependency 'caxlsx', ['>= 3.3.0', '<4']
  s.add_runtime_dependency 'rodf', ['>= 1.0.0', '<2']

  s.add_development_dependency 'rake'
  s.add_development_dependency 'minitest'
  s.add_development_dependency 'minitest-reporters'
  s.add_development_dependency 'pry'

  if RUBY_VERSION.to_f >= 2.4
    s.add_development_dependency 'warning'
  end
end
