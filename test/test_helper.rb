ENV["RAILS_ENV"] = "test"

require 'pry'

begin
  require 'warning'

  Warning.ignore(
    %r{mail/parsers/address_lists_parser}, ### Hide mail gem warnings
  )
rescue LoadError
  # Do nothing
end

require 'pry'

require File.expand_path("../dummy_app/config/environment.rb",  __FILE__)

migration_path = Rails.root.join('db/migrate')
if ActiveRecord::VERSION::MAJOR == 6
  ActiveRecord::MigrationContext.new(File.expand_path(migration_path, __dir__), ActiveRecord::SchemaMigration).migrate
else
  ActiveRecord::MigrationContext.new(File.expand_path(migration_path, __dir__)).migrate
end

require "rails/test_help"

Rails.backtrace_cleaner.remove_silencers!

class ActiveSupport::TestCase
  # Setup all fixtures in test/fixtures/*.yml for all tests in alphabetical order.
  fixtures :all
end

require 'minitest-spec-rails' ### for describe blocks

require 'minitest/reporters'
Minitest::Reporters.use!(
  Minitest::Reporters::DefaultReporter.new,
  ENV,
  Minitest.backtrace_filter
)

require 'minitest-spec-rails'

post_count = Post.count
if post_count < 5
  (5 - post_count).times do |i|
    Post.create!(name: "foo #{i}", content: "bar #{i}", age: i)
  end
end

TMP_PATH = Rails.root.join("../../tmp/")

### Cleanup old test spreadsheets
FileUtils.remove_dir(TMP_PATH, true)
FileUtils.mkdir_p(TMP_PATH)

def save_file(path, file_data)
  if file_data.is_a?(Axlsx::Package)
    file_data = file_data.to_stream.read
  elsif file_data.is_a?(RODF::Spreadsheet)
    file_data = file_data.bytes
  end

  path = Rails.root.join("../../tmp/", path)

  FileUtils.mkdir_p(File.dirname(path))

  File.open(TMP_PATH.join(path), "w+b") do |f|
    f.write file_data
  end
end

def parse_ods_spreadsheet(spreadsheet)
  Nokogiri::XML(spreadsheet.xml)
end

def parse_axlsx_package(package)
  Nokogiri::XML(package.workbook.worksheets.first.to_xml_string)
end

def parse_axlsx_worksheet(worksheet)
  Nokogiri::XML(worksheet.to_xml_string)
end
