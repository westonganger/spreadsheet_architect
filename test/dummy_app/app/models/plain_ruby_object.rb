class PlainRubyObject
  def spreadsheet_columns
    [
      ['Title', :title],
      :content,
      ['Object ID', (object_id)]
    ]
  end

  def title
    'My Title'
  end

  def content
    'The content...'
  end
end
