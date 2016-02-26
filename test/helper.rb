config = YAML::load(IO.read(File.dirname(__FILE__) + '/database.yml'))
ActiveRecord::Base.establish_connection(config['sqlite3'])
ActiveRecord::Schema.define(version: 0) do
  begin
  drop_table :posts, :force => true
  drop_table :other_posts, :force => true
  rescue
    #dont really care if the tables are not dropped
  end

  create_table(:posts, :force => true) do |t|
    t.string :name
    t.string :title
    t.text :content
    t.integer :votes
    t.timestamps null: false
  end

  create_table(:other_posts, :force => true) do |t|
    t.string :name
    t.string :title
    t.text :content
    t.integer :votes
    t.timestamps null: false
  end
end

class Post < ActiveRecord::Base
  include SpreadsheetArchitect

  def spreadsheet_columns
    [:name, :title, :content, :votes, :ranking]
  end

  def ranking
    1
  end
end

class OtherPost < ActiveRecord::Base
  include SpreadsheetArchitect
end

posts = []
posts << Post.new(name: "first post", title: "This is the first post", content: "I am a very good first post!", votes: 1)
posts << Post.new(name: "second post", title: "This is the second post", content: "I am the best post!", votes: 7)
posts.each { |p| p.save! }

posts = []
posts << OtherPost.new(name: "my other first", title: "first other post", content: "the first other post!", votes: 1)
posts << OtherPost.new(name: "my other second", title: "second other post", content: "last other post!", votes: 7)
posts.each { |p| p.save! }
