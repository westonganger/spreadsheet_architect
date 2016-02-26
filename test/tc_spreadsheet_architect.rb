#!/usr/bin/env ruby -w
require "spreadsheet_architect"
require 'yaml'
require 'active_record'
require 'minitest'

require File.expand_path(File.join(File.dirname(__FILE__), 'helper'))

class TestSpreadsheetArchitect < MiniTest::Test
  class Post < ActiveRecord::Base
    include SpreadsheetArchitect

    def self.spreadsheet_columns
      [:name, :title, :content, :votes, :ranking]
    end

    def ranking
      1
    end
  end

  class OtherPost < ActiveRecord::Base
    include SpreadsheetArchitect
  end

  class PlainPost
    include SpreadsheetArchitect

    def self.spreadsheet_columns
      [:name, :title, :content]
    end

    def name
      "the name"
    end

    def title
      "the title"
    end

    def content
      "the content"
    end
  end

  test "test_spreadsheet_options" do
    assert_equal([:name, :title, :content, :votes, :ranking], Post.spreadsheet_columns)
    assert_equal([:name, :title, :content, :votes, :created_at, :updated_at], OtherPost.column_names)
    assert_equal([:name, :title, :content], PlainPost.spreadsheet_columns)
  end
end
  
class TestToCsv < MiniTest::Test
  test "test_class_method" do
    p = Post.to_csv(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
  test 'test_chained_method' do
    p = Post.order("name asc").to_csv(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
end

class TestToOds < MiniTest::Test
  test 'test_class_method' do
    p = Post.to_ods(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
  test 'test_chained_method' do
    p = Post.order("name asc").to_ods(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
end

class TestToXlsx < MiniTest::Test
  test 'test_class_method' do
    p = Post.to_xlsx(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
  test 'test_chained_method' do
    p = Post.order("name asc").to_xlsx(spreadsheet_columns: [:name, :votes, :content, :ranking])
    assert_equal(true, p.is_a?(String))
  end
end
