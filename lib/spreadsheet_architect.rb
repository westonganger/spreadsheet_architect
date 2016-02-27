require 'spreadsheet_architect/set_mime_types'
require 'spreadsheet_architect/action_controller_renderers'
require 'axlsx'
require 'spreadsheet_architect/axlsx_column_width_patch'
require 'odf/spreadsheet'
require 'csv'

module SpreadsheetArchitect
  def self.included(base)
    base.send :extend, ClassMethods
  end

  module ClassMethods
    def sa_str_humanize(str, capitalize = true)
      str = str.sub(/\A_+/, '').gsub(/[_\.]/,' ').sub(' rescue nil','')
      if capitalize
        str = str.gsub(/(\A|\ )\w/){|x| x.upcase}
      end
      str
    end

    def sa_get_options(options={})
      if self.ancestors.include?(ActiveRecord::Base) && !self.respond_to?(:spreadsheet_columns) && !options[:spreadsheet_columns]
        the_column_names = (self.column_names - ["id","created_at","updated_at","deleted_at"])
        headers = the_column_names.map{|x| x.humanize}
        columns = the_column_names.map{|x| x.to_s}
      elsif options[:spreadsheet_columns] || self.respond_to?(:spreadsheet_columns)
        headers = []
        columns = []

        array = options[:spreadsheet_columns] || self.spreadsheet_columns || []
        array.each do |x|
          if x.is_a?(Array)
            headers.push x[0].to_s
            columns.push x[1].to_s
          else
            headers.push sa_str_humanize(x.to_s)
            columns.push x.to_s
          end
        end
      else
        headers = []
        columns = []
      end

      headers = (options[:headers] == false ? false : headers)

      header_style = {background_color: "AAAAAA", color: "FFFFFF", align: :center, bold: false, font_name: 'Arial', font_size: 10, italic: false, underline: false}
      
      if options[:header_style]
        header_style.merge!(options[:header_style])
      elsif options[:header_style] == false
        header_style = false
      elsif options[:row_style]
        header_style = options[:row_style]
      end

      row_style = {background_color: nil, color: "000000", align: :left, bold: false, font_name: 'Arial', font_size: 10, italic: false, underline: false}
      if options[:row_style]
        row_style.merge!(options[:row_style])
      end

      sheet_name = options[:sheet_name] || self.name
      
      if options[:data]
        data = options[:data].to_a
      elsif self.ancestors.include?(ActiveRecord::Base)
        data = where(options[:where]).order(options[:order]).to_a
      else
        # object must have a to_a method 
        data = self.to_a
      end

      types = (options[:types] || []).flatten

      return {headers: headers, columns: columns, header_style: header_style, row_style: row_style, types: types, sheet_name: sheet_name, data: data}
    end

    def sa_get_row_data(the_columns=[], instance)
      row_data = []
      the_columns.each do |col|
        col.split('.').each_with_index do |x,i| 
          if i == 0
            col = instance.instance_eval(x)
          else
            col = col.instance_eval(x)
          end
        end
        row_data.push col.to_s
      end
      return row_data
    end

    def to_csv(opts={})
      options = sa_get_options(opts)

      CSV.generate do |csv|
        csv << options[:headers] if options[:headers]
        
        options[:data].each do |x|
          csv << sa_get_row_data(options[:columns], x)
        end
      end
    end

    def to_ods(opts={})
      options = sa_get_options(opts)

      spreadsheet = ODF::Spreadsheet.new

      spreadsheet.office_style :header_style, family: :cell do
        if options[:header_style]
          unless opts[:row_style] && opts[:row_style][:bold] == false #uses opts, temporary
            property :text, 'font-weight': :bold
          end
          if options[:header_style][:align]
            property :text, 'align': options[:header_style][:align]
          end
          if options[:header_style][:size]
            property :text, 'font-size': options[:header_style][:size]
          end
          if options[:header_style][:color] && opts[:header_style] && opts[:header_style][:color] #temporary
            property :text, 'color': "##{options[:header_style][:color]}"
          end
        end
      end
      spreadsheet.office_style :row_style, family: :cell do
        if options[:row_style]
          if options[:row_style][:bold]
            property :text, 'font-weight': :bold
          end
          if options[:row_style][:align]
            property :text, 'align': options[:row_style][:align]
          end
          if options[:row_style][:size]
            property :text, 'font-size': options[:row_style][:size]
          end
          if opts[:row_style] && opts[:row_style][:color] #uses opts, temporary
            property :text, 'color': "##{options[:row_style][:color]}"
          end
        end
      end

      this_class = self
      spreadsheet.table options[:sheet_name] do 
        if options[:headers]
          row do
            options[:headers].each do |header|
              cell header, style: :header_style
            end
          end
        end
        options[:data].each do |x|
          row do 
            this_class.sa_get_row_data(options[:columns], x).each do |y|
              cell y, style: :row_style
            end
          end
        end
      end

      return spreadsheet.bytes
    end

    def to_xlsx(opts={})
      options = sa_get_options(opts)
    
      header_style = {}
      header_style[:fg_color] = options[:header_style].delete(:color)
      header_style[:bg_color] = options[:header_style].delete(:background_color)
      if header_style[:align]
        header_style[:alignment] = {}
        header_style[:alignment][:horizontal] = options[:header_style][:align]
      end
      header_style[:b] = options[:header_style].delete(:bold)
      header_style[:sz] = options[:header_style].delete(:font_size)
      header_style[:i] = options[:header_style].delete(:italic)
      header_style[:u] = options[:header_style].delete(:underline)
      
      row_style = {}
      row_style[:fg_color] = options[:row_style].delete(:color)
      row_style[:bg_color] = options[:row_style].delete(:background_color)
      if row_style[:align]
        row_style[:alignment] = {}
        row_style[:alignment][:horizontal] = options[:row_style][:align]
      end
      row_style[:b] = options[:row_style].delete(:bold)
      row_style[:sz] = options[:row_style].delete(:font_size)
      row_style[:i] = options[:row_style].delete(:italic)
      row_style[:u] = options[:row_style].delete(:underline)
      
      package = Axlsx::Package.new

      return package if options[:data].empty?

      package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|
        if options[:headers]
          sheet.add_row options[:headers], style: package.workbook.styles.add_style(header_style)
        end
        
        options[:data].each do |x|
          sheet.add_row sa_get_row_data(options[:columns], x), style: package.workbook.styles.add_style(row_style), types: options[:types]
        end
      end

      return package.to_stream.read
    end
  end
end
