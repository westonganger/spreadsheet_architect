class SpreadsheetsController < ApplicationController

  respond_to :csv, :xlsx, :ods
  
  def csv
    render csv: Post.to_csv
  end

  def ods
    render ods: Post.to_ods
  end

  def xlsx
    render xlsx: Post.to_xlsx, filename: 'Posts'
  end

  def alt_xlsx
    @posts = Post.all
    respond_with @posts
  end

end
