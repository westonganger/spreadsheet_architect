Dummy::Application.routes.draw do
  get 'spreadsheet/csv', to: 'spreadsheets#csv'
  get 'spreadsheet/ods', to: 'spreadsheets#ods'
  get 'spreadsheet/xlsx', to: 'spreadsheets#xlsx'
  get 'spreadsheet/alt_xlsx', to: 'spreadsheets#alt_xlsx'
end
