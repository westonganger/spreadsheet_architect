if defined? Mime
  unless defined? Mime::XLSX
    Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx
  end
  unless defined? Mime::ODS
    Mime::Type.register "application/vnd.oasis.opendocument.spreadsheet", :ods
  end
end
