require 'test_helper'

class OdsTest < ActiveSupport::TestCase

  def setup
    FileUtils.mkdir_p(Rails.root.join('tmp/ods'))

    10.times.each do
      CustomPost.create(name: SecureRandom.hex, content: SecureRandom.hex, age: 2)
    end

    @headers = ['test1', 'test2','test3']
    @data = [
      ['row1','test1'],
      ['row2 c1', nil, '123'],
      ['the',Date.today,Time.now]
    ]
    @options = {
      headers: true,
      header_style: {background_color: 'AAAAAA', color: 'FFFFFF', align: :center, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      row_style: {background_color: nil, color: '000000', align: :left, font_name: 'Arial', font_size: 10, bold: false, italic: false, underline: false},
      sheet_name: 'My Project Export',
      column_styles: [],
      range_styles: [],
      merges: [],
      borders: [],
      column_types: [:string, :date, :date]
    }
  end

  def test_empty_model
    CustomPost.delete_all
    File.open(Rails.root.join('tmp/ods/empty_model.ods'),'w+b') do |f|
      f.write CustomPost.to_ods
    end
  end

  def test_empty_sa
    File.open(Rails.root.join('tmp/ods/empty_sa.ods'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_ods(data: [])
    end
  end

  def test_sa
    File.open(Rails.root.join('tmp/ods/sa.ods'),'w+b') do |f|
      f.write SpreadsheetArchitect.to_ods(headers: @headers, data: @data)
    end
  end

  def test_model
    File.open(Rails.root.join('tmp/ods/model.ods'),'w+b') do |f|
      f.write CustomPost.to_ods
    end
  end

  def test_options
    File.open(Rails.root.join('tmp/ods/model_options.ods'),'w+b') do |f|
      f.write CustomPost.to_ods(@options)
    end
  end

  def teardown
  end

end
