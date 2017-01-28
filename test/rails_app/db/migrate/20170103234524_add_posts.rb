class AddPosts < ActiveRecord::Migration[5.0]
  def change
    create_table :posts do |t|
      t.string :name
      t.text :content
      t.integer :age
      t.timestamps
    end
  end
end
