require "test/test_helper"

class PostTest < ActiveSupport::TestCase

  def setup
    @path = 'tmp/posts'
    FileUtils.mkdir_p(@path)

    DatabaseCleaner.start
  end

  def test_csv
    data = Post.to_csv

    File.open(File.join(@path, 'csv.csv'), 'w+b') do |f|
      f.write data
    end
  end

  def test_ods
    data = Post.to_ods

    File.open(File.join(@path, 'ods.ods'), 'w+b') do |f|
      f.write data
    end
  end

  def test_xlsx
    data = Post.all.to_xlsx

    File.open(File.join(@path, 'xlsx.xlsx'), 'w+b') do |f|
      f.write data
    end
  end

  def test_empty
    Post.destroy_all
    
    File.open(File.join(@path, 'empty.xlsx'), 'w+b') do |f|
      f.write Post.to_xlsx
    end
  end

  def teardown
    DatabaseCleaner.clean
  end
end
