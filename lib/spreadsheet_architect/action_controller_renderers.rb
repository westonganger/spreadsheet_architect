if defined? ActionController
  ActionController::Renderers.add :xlsx do |data, options|
    if data.is_a?(ActiveRecord::Relation)
      options[:filename] = data.model.name.pluralize
      data = data.to_xlsx
    end
    send_data data, type: :xlsx, disposition: :attachment, filename: "#{options[:filename] ? options[:filename].sub('.xlsx','') : 'data'}.xlsx"
  end
  ActionController::Renderers.add :ods do |data, options|
    if data.is_a?(ActiveRecord::Relation)
      options[:filename] = data.model.name.pluralize
      data = data.to_ods
    end
    send_data data, type: :ods, disposition: :attachment, filename: "#{options[:filename] ? options[:filename].sub('.ods','') : 'data'}.ods"
  end
  ActionController::Renderers.add :csv do |data, options|
    if data.is_a?(ActiveRecord::Relation)
      options[:filename] = data.model.name.pluralize
      data = data.to_csv
    end
    send_data data, type: :csv, disposition: :attachment, filename: "#{options[:filename] ? options[:filename].sub('.csv','') : 'data'}.csv"
  end
end
