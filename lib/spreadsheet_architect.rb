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
      str = str.sub(/\A_+/, '').sub(/_id\z/, '').gsub(/[_\.]/,' ').sub(' rescue nil','')
      if capitalize
        str = str.gsub(/(\A|\ )\w/){|x| x.upcase}
      end
      str
    end

    def sa_get_options(options={})
      if self.ancestors.include?(ActiveRecord::Base) && !self.respond_to?(:spreadsheet_columns) && !options[:spreadsheet_columns]
        headers = self.column_names.map{|x| x.humanize}
        columns = self.column_names.map{|x| x.to_s}
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

      if options[:header_style]
        header_style = options[:header_style]
      elsif options[:header_style] == false
        header_style = false
      elsif options[:row_style]
        header_style = options[:row_style]
      else
        header_style = {bg_color: "AAAAAA", fg_color: "FFFFFF", alignment: { horizontal: :center }, bold: true}
      end

      row_style = options[:row_style]

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
        row_data.push col
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
          if options[:header_style][:bold]
            property :text, 'font-weight': :bold
            property :text, 'align': :center
          end
          if options[:header_style][:fg_color] && opts[:header_style] && opts[:header_style][:fg_color] #temporary
            property :text, 'color': "##{options[:header_style][:fg_color]}"
          end
        end
      end
      spreadsheet.office_style :row_style, family: :cell do
        if options[:row_style]
          if options[:row_style][:bold]
            property :text, 'font-weight': :bold
          end
          if options[:row_style][:fg_color]
            property :text, 'color': "##{options[:header_style][:fg_color]}"
          end
        end
      end

      this_class = self
      spreadsheet.table options[:sheet_name] do 
        if options[:headers]
          row do
            options[:headers].each do |header|
              cell header, style: (:header_style if options[:header_style])
            end
          end
        end
        options[:data].each do |x|
          row do 
            this_class.sa_get_row_data(options[:columns], x).each do |y|
              cell y, style: (:row_style if options[:row_style])
            end
          end
        end
      end

      return spreadsheet.bytes
    end

    def to_xlsx(opts={})
      options = sa_get_options(opts)
      
      package = opts[:package] || Axlsx::Package.new

      return package if options[:data].empty?

      package.workbook.add_worksheet(name: options[:sheet_name]) do |sheet|
        if options[:headers]
          sheet.add_row options[:headers], style: (package.workbook.styles.add_style(options[:header_style]) if options[:header_style])
        end
        
        options[:data].each do |x|
          sheet.add_row sa_get_row_data(options[:columns], x), style: (package.workbook.styles.add_style(options[:row_style]) if options[:row_style]), types: options[:types]
        end
      end

      return package.to_stream.read
    end
  end
end
