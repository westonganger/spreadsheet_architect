require File.expand_path(File.dirname(__FILE__) + '/lib/spreadsheet_architect/version.rb')
require 'bundler/gem_tasks'

task :test do 
  require 'rake/testtask'
  Rake::TestTask.new do |t|
    t.libs << 'test'
    t.test_files = FileList['test/**/tc_*.rb']
    t.verbose = true
  end
end

task release: :build do
  system "gem push spreadsheet_architect-#{SpreadsheetArchitect::VERSION}.gem"
end

task default: :test
