if defined? ActionController

  ['csv','ods','xlsx'].each do |format|
    ActionController::Renderers.add(format.to_sym) do |data, options|
      if defined?(ActiveRecord) && data.is_a?(ActiveRecord::Relation)
        options[:filename] ||= data.klass.name.pluralize
        data = data.send("to_#{format}")
      end

      options[:filename] = options[:filename] ? options[:filename].strip.sub(/\.#{format}$/i,'') : 'data'
      options[:filename] += ".#{format}"
      
      send_data data, type: format.to_sym, disposition: :attachment, filename: options[:filename]
    end
  end

end
